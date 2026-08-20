// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// PoC: resolve a Teams meeting join URL and fetch its transcript via Microsoft Graph.
//
// Run:
//   npm start -- <userId> "<teams-meeting-url>"
//
// Environment (.env):
//   TENANT_ID, CLIENT_ID, CLIENT_SECRET  - app registration credentials
//
// App permissions (admin consent):
//   OnlineMeetings.Read.All, OnlineMeetingTranscript.Read.All

import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/teams.graph';
import * as betaEndpoints from '@microsoft/teams.graph-endpoints-beta';
import * as dotenv from 'dotenv';

dotenv.config();

// --- common ---

const credential = new ClientSecretCredential(
  process.env.TENANT_ID || '',
  process.env.CLIENT_ID || '',
  process.env.CLIENT_SECRET || ''
);

const graphClient = new Client({
  token: async () => {
    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
    return tokenResponse.token;
  },
});

// --- getMeetingIdByUrl ---

interface ParsedMeetingUrl {
  joinUrl: string;
  threadId: string;
  organizerUserId?: string;
  tenantId?: string;
}

interface OnlineMeeting {
  id?: string;
  joinWebUrl?: string;
  subject?: string;
}

interface GraphListResponse<T> {
  value?: T[];
}

interface MeetingLookupResult {
  meetingId: string;
  userId: string;
}

function normalizeJoinUrl(raw: string): string {
  return raw.trim().replace(/\\/g, '');
}

function joinUrlVariants(raw: string): string[] {
  const normalized = normalizeJoinUrl(raw);
  const url = new URL(normalized);
  const variants = new Set<string>();

  const singleEncoded = `${url.origin}${url.pathname}${url.search}`;
  const decodedPath = decodeURIComponent(url.pathname);

  // App-only tokens require double-encoded JoinWebUrl in the filter (Graph docs, Example 3).
  variants.add(encodeURIComponent(singleEncoded));
  variants.add(encodeURIComponent(url.href));

  variants.add(normalized);
  variants.add(url.href);
  variants.add(singleEncoded);

  // Decoded path + query (delegated-token format in Graph docs).
  variants.add(`${url.origin}${decodedPath}${url.search}`);

  const context = url.searchParams.get('context');
  if (context) {
    variants.add(`${url.origin}${decodedPath}?context=${context}`);
    try {
      variants.add(`${url.origin}${decodedPath}?context=${JSON.stringify(JSON.parse(context))}`);
    } catch {
      // ignore malformed context JSON
    }
  }

  variants.add(`${url.origin}${url.pathname}`);
  variants.add(`${url.origin}${decodedPath}`);

  return [...variants];
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function resolveOnlineMeeting(userId: string, joinUrl: string): Promise<OnlineMeeting | undefined> {
  for (const candidate of joinUrlVariants(joinUrl)) {
    try {
      const filter = `JoinWebUrl eq '${escapeODataString(candidate)}'`;
      const response = (
        await graphClient.http.get(`/users/${userId}/onlineMeetings`, {
          params: { $filter: filter },
        })
      ).data as GraphListResponse<OnlineMeeting>;

      console.log("meeting - Graph API response: ");
      console.log(response);

      const meeting = response.value?.[0];
      if (meeting?.id) {
        return meeting;
      }
    } catch {
      // Try the next URL encoding variant (e.g. app-only needs double-encoded JoinWebUrl).
    }
  }

  return undefined;
}

async function getMeetingIdByUrl(userId: string, joinUrl: string): Promise<MeetingLookupResult> {
  const meeting = await resolveOnlineMeeting(userId, joinUrl);

  if (!meeting?.id) {
    throw new Error(
      `No online meeting found for join URL (organizer: ${userId})`
    );
  }

  return { meetingId: meeting.id, userId };
}

// --- getMeetingTranscript ---

interface MeetingTranscriptResult {
  id: string;
  content: string;
}

const emptyTranscript: MeetingTranscriptResult = { id: '', content: '' };

async function getMeetingTranscript(
  meetingResourceId: string,
  userId: string
): Promise<MeetingTranscriptResult> {
  try {
    const transcriptsResponse = await graphClient.call(
      betaEndpoints.users.onlineMeetings.transcripts.list,
      {
        'user-id': userId,
        'onlineMeeting-id': meetingResourceId,
      }
    );

    if (!transcriptsResponse.value?.length) {
      return emptyTranscript;
    }

    const latestTranscript = transcriptsResponse.value.reduce((latest, current) => {
      const latestDate = latest.createdDateTime ? new Date(latest.createdDateTime) : new Date(0);
      const currentDate = current.createdDateTime ? new Date(current.createdDateTime) : new Date(0);
      return currentDate > latestDate ? current : latest;
    }, transcriptsResponse.value[0]);

    if (!latestTranscript.id) {
      return emptyTranscript;
    }

    const content = await graphClient.call(
      betaEndpoints.users.onlineMeetings.transcripts.content.get,
      {
        'user-id': userId,
        'onlineMeeting-id': meetingResourceId,
        'callTranscript-id': latestTranscript.id,
      },
      { requestConfig: { headers: { Accept: 'text/vtt' } } }
    );

    return { id: latestTranscript.id, content: content ?? '' };
  } catch (error) {
    console.error('Error retrieving transcript:', formatGraphApiError(error));
    return emptyTranscript;
  }
}

// --- main ---

function formatGraphApiError(error: unknown): string {
  if (error && typeof error === 'object') {
    const response = (error as { response?: { status?: number; data?: unknown } }).response;
    if (response) {
      const status = response.status ?? 'unknown';
      const body =
        response.data === undefined
          ? ''
          : typeof response.data === 'string'
            ? response.data
            : JSON.stringify(response.data, null, 2);
      const message = error instanceof Error ? error.message : 'Graph API request failed';
      return body ? `${message} (${status})\n${body}` : `${message} (${status})`;
    }
  }

  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function usage(): never {
  console.error('Usage: npm start -- <userId> <teams-meeting-url>');
  process.exit(1);
}

function validateEnv(): void {
  const missing: string[] = [];

  if (!process.env.TENANT_ID?.trim()) missing.push('TENANT_ID');
  if (!process.env.CLIENT_ID?.trim()) missing.push('CLIENT_ID');
  if (!process.env.CLIENT_SECRET?.trim()) missing.push('CLIENT_SECRET');

  if (missing.length > 0) {
    console.error(`Missing required environment variable(s): ${missing.join(', ')}`);
    process.exit(1);
  }
}

async function main(): Promise<void> {
  validateEnv();
  const userId = process.argv[2];
  const joinUrl = process.argv[3];
  if (!userId || userId.startsWith('-') || !joinUrl || joinUrl.startsWith('-')) {
    usage();
  }

  const { meetingId } = await getMeetingIdByUrl(userId.trim(), joinUrl.trim());
  const transcript = await getMeetingTranscript(meetingId, userId);
  console.log(`transcript id: ${transcript.id}`);
  console.log(transcript.content || '(no transcript available)');
}

main().catch((error) => {
  console.error(formatGraphApiError(error));
  process.exit(1);
});
