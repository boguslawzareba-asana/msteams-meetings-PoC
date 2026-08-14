// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// PoC: resolve a Teams meeting join URL and fetch its transcript via Microsoft Graph.
//
// Run:
//   npm start -- "<teams-meeting-url>"
//
// Environment (.env):
//   TENANT_ID, CLIENT_ID, CLIENT_SECRET  - app registration credentials
//   TARGET_USER_ID or MEETING_ORGANIZER_USER_ID - fallback organizer id when not in URL context
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

function joinUrlVariants(raw: string): string[] {
  const url = new URL(raw);
  const variants = new Set<string>();

  variants.add(raw.trim());
  variants.add(url.toString());

  const withoutQuery = `${url.origin}${url.pathname}`;
  variants.add(withoutQuery);

  const decodedPath = decodeURIComponent(url.pathname);
  variants.add(`${url.origin}${decodedPath}`);

  return [...variants];
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function resolveOnlineMeeting(userId: string, joinUrl: string): Promise<OnlineMeeting | undefined> {
  for (const candidate of joinUrlVariants(joinUrl)) {
    const filter = `JoinWebUrl eq '${escapeODataString(candidate)}'`;
    const response = (
      await graphClient.http.get(`/users/${userId}/onlineMeetings`, {
        params: { $filter: filter },
      })
    ).data as GraphListResponse<OnlineMeeting>;

    const meeting = response.value?.[0];
    if (meeting?.id) {
      return meeting;
    }
  }

  return undefined;
}

async function getMeetingIdByUrl(joinUrl: string): Promise<MeetingLookupResult> {
  const userId = process.env.TARGET_USER_ID;
  const meeting = await resolveOnlineMeeting(userId, joinUrl);

  if (!meeting?.id) {
    throw new Error(
      `No online meeting found for join URL (organizer: ${userId})`
    );
  }

  return { meetingId: meeting.id, userId };
}

// --- getMeetingTranscript ---

async function getMeetingTranscript(meetingResourceId: string, userId: string): Promise<string> {
  try {
    const transcriptsResponse = await graphClient.call(
      betaEndpoints.users.onlineMeetings.transcripts.list,
      {
        'user-id': userId,
        'onlineMeeting-id': meetingResourceId,
      }
    );

    if (!transcriptsResponse.value?.length) {
      return '';
    }

    const latestTranscript = transcriptsResponse.value.reduce((latest, current) => {
      const latestDate = latest.createdDateTime ? new Date(latest.createdDateTime) : new Date(0);
      const currentDate = current.createdDateTime ? new Date(current.createdDateTime) : new Date(0);
      return currentDate > latestDate ? current : latest;
    }, transcriptsResponse.value[0]);

    if (!latestTranscript.id) {
      return '';
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

    return content ?? '';
  } catch (error) {
    console.error('Error retrieving transcript:', error);
    return '';
  }
}

// --- main ---

function usage(): never {
  console.error('Usage: npm start -- <teams-meeting-url>');
  process.exit(1);
}

function validateEnv(): void {
  const missing: string[] = [];

  if (!process.env.TENANT_ID?.trim()) missing.push('TENANT_ID');
  if (!process.env.CLIENT_ID?.trim()) missing.push('CLIENT_ID');
  if (!process.env.CLIENT_SECRET?.trim()) missing.push('CLIENT_SECRET');
  if (!process.env.TARGET_USER_ID?.trim()) missing.push('TARGET_USER_ID');

  if (missing.length > 0) {
    console.error(`Missing required environment variable(s): ${missing.join(', ')}`);
    process.exit(1);
  }
}

async function main(): Promise<void> {
  validateEnv();
  const joinUrl = process.argv[2];
  if (!joinUrl || joinUrl.startsWith('-')) {
    usage();
  }

  const { meetingId, userId } = await getMeetingIdByUrl(joinUrl.trim());
  const transcript = await getMeetingTranscript(meetingId, userId);
  console.log(transcript || '(no transcript available)');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
