type Person = {
  Id?: number;
  Title?: string;
};
export type attachments = {
  FileName?: string;
  ServerRelativeUrl?: string;
};
export interface ISPListItem {
  ID?: number;
  Title: string;
  Note?: string;
  Assignee: Person;
  AttachmentFiles?: attachments[];
  DueDate?: Date;
}
