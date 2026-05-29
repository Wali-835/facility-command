export interface Contact {
  id: string;
  firstName: string;
  lastName: string;
  email: string;
  company: string;
  jobTitle: string;
  phone: string;
  tags: string[];
  addedAt: string;
}

export type CampaignStatus = 'draft' | 'sent';

export interface Campaign {
  id: string;
  name: string;
  templateId: string;
  contactIds: string[];
  status: CampaignStatus;
  senderName: string;
  senderEmail: string;
  customSubject: string;
  createdAt: string;
  sentAt?: string;
}

export interface EmailTemplate {
  id: string;
  name: string;
  subject: string;
  body: string;
  category: 'introduction' | 'roi' | 'demo' | 'followup';
}

export interface LinkedInPost {
  id: string;
  postType: string;
  machine: string;
  tone: string;
  content: string;
  savedAt: string;
}
