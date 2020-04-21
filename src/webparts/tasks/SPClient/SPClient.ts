import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { ISPListItem } from "./ISPListItem";

import { BaseWebPartContext } from "@microsoft/sp-webpart-base";

export default class SPClient {
  constructor(public context: BaseWebPartContext) {}
  public getRecentItems(): Promise<ISPListItem[]> {
    const restUrl: string =
      this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists('A4F01117-210D-46C5-91C1-CCA4CCDA57A1')/items?$top=3&$expand=Assignee,AttachmentFiles&$select=AttachmentFiles,ID,Title,Assignee/Id,Assignee/Title,Note,DueDate&$orderBy=Created desc";
    return this.context.spHttpClient
      .get(restUrl, SPHttpClient.configurations.v1)
      .then((data: SPHttpClientResponse) => data.json())
      .then((data) => data.value);
  }

  public getItem(id: number): Promise<ISPListItem> {
    const restUrl: string =
      this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists('A4F01117-210D-46C5-91C1-CCA4CCDA57A1')/items(" +
      id +
      ")?$top=3&$expand=Assignee,AttachmentFiles&$select=AttachmentFiles,ID,Title,Assignee/Id,Assignee/Title,Note,DueDate";
    return this.context.spHttpClient
      .get(restUrl, SPHttpClient.configurations.v1)
      .then((data: SPHttpClientResponse) => data.json())
      .then((data) => data.value);
  }

  public createItem(item: ISPListItem): Promise<number> {
    let newTask: any = {
      Title: item.Title,
      Note: item.Note,
      DueDate: item.DueDate,
      AssigneeId: item.Assignee.Id,
    };
    const restUrl =
      this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists('A4F01117-210D-46C5-91C1-CCA4CCDA57A1')/items";
    let _SPHTTPClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(newTask),
    };
    return this.context.spHttpClient
      .post(restUrl, SPHttpClient.configurations.v1, _SPHTTPClientOptions)
      .then((data: SPHttpClientResponse) => data.json())
      .then((data) => data.ID) as Promise<number>;
  }
  private _uploadFile(file: any): Promise<any> {
    const reader = new FileReader();
    return new Promise((resolve, reject) => {
      reader.onerror = (e) => reject(e.target.error);
      reader.onload = (e) => resolve(e.target.result);
      reader.readAsArrayBuffer(file);
    });
  }
  public async uploadAttachments(files: FileList, id: number) {
    for (let i = 0; i < files.length; i++) {
      let restUrl: string =
        this.context.pageContext.web.absoluteUrl +
        "/_api/web/lists('A4F01117-210D-46C5-91C1-CCA4CCDA57A1')/items(" +
        id +
        ")/AttachmentFiles/add(FileName='" +
        files[i].name +
        "')";
      await this._uploadFile(files[i])
        .then((data) => {
          let options: ISPHttpClientOptions = { body: data };
          return this.context.spHttpClient.post(
            restUrl,
            SPHttpClient.configurations.v1,
            options
          );
        })
        .then((result) => result.json());
    }
  }
}
