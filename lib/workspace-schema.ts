import { z } from 'zod';
const text = z.string().trim().max(2000);
const connection = z.object({ account: text, requested: z.boolean() }).strict();
export const workspaceSchema = z.object({
  company: text, contact: text, email: z.string().email().or(z.literal('')),
  industry: text, volume: text, mailboxes: text, provider: text,
  goals: z.array(z.string().max(100)).max(10), pain: text, urgency: text,
  languages: text, attachments: text, access: text, notes: text,
  submitted: z.boolean(), consent: z.boolean(),
  connections: z.object({ outlook: connection, google: connection, telegram: connection }).strict(),
}).strict().refine(data => !data.submitted || (data.consent && data.company && data.contact && data.email), {
  message: 'Complete your contact details and consent before submitting.',
});
export const workspaceWriteSchema = z.object({
  data: workspaceSchema,
  workspaceId: z.string().regex(/^[a-f0-9]{64}$/).nullable(),
  revision: z.number().int().min(1).max(2147483646).nullable(),
}).strict();
export type Workspace = z.infer<typeof workspaceSchema>;
