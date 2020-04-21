import { ISPListItem } from "./ISPListItem";
export default class MockSPClient {
  private static _listItems: ISPListItem[] = [
    {
      ID: 1,
      Title: "Mock SharePoint Client",
      Assignee: { Id: 12, Title: "Mehdi Mowlavi" },
      Note: "This is a test",
      DueDate: new Date("2019-03-03"),
      AttachmentFiles: [
        { FileName: "Dev.pptx", ServerRelativeUrl: "/spfx.pptx" },
        { FileName: "Dev.docx", ServerRelativeUrl: "/spfx.docx" },
      ],
    },
    {
      ID: 2,
      Title: "Make React Components",
      Assignee: { Id: 5, Title: "Mehdi H" },
      Note: "This is another test",
      DueDate: new Date("2020-01-12"),
      AttachmentFiles: [
        { FileName: "Dev.pptx", ServerRelativeUrl: "/spfx.pptx" },
      ],
    },
    {
      ID: 3,
      Title: "Integrate Things",
      Assignee: { Id: 4, Title: "Ehsan M" },
      Note: "This is another test",
      DueDate: new Date("2020-08-08"),
      AttachmentFiles: [],
    },
  ];
  public static get(restUrl: string, options?: any): Promise<ISPListItem[]> {
    return new Promise<ISPListItem[]>((resolve, reject) => {
      resolve(MockSPClient._listItems);
    });
  }
}
