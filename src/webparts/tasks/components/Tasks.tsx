import * as React from "react";
import styles from "./Tasks.module.scss";
import { ITasksProps } from "./ITasksProps";
import ITasksState from "./ITasksState";
import { ISPListItem } from "../SPClient/ISPListItem";
import MockSPClient from "../SPClient/MockSPClient";
import SPClient from "../SPClient/SPClient";

import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

import { DetailsList, IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { Link } from "office-ui-fabric-react/lib/Link";

import NewTask from "./forms/newform";
import { attachments } from "../SPClient/ISPListItem";

const selectedColumns: IColumn[] = [
  {
    name: "Title",
    fieldName: "Title",
    key: "title",
    minWidth: 50,
    headerClassName: "ListHeaders",
    isResizable: true,
    maxWidth: 200,
  },
  {
    name: "Assigned To",
    key: "assignee",
    fieldName: "Assignee",
    minWidth: 50,
    maxWidth: 100,
    headerClassName: "ListHeaders",
    onRender: (item, index, column) => {
      if (item.Assignee) return item.Assignee.Title;
      else return "";
    },
  },
  {
    name: "Due Date",
    key: "dueDate",
    fieldName: "DueDate",
    headerClassName: "ListHeaders",
    minWidth: 50,
    maxWidth: 100,
    onRender: (item) => {
      if (item.DueDate) return new Date(item.DueDate).toDateString();
      else return "";
    },
  },
  {
    name: "Attachments",
    key: "attachments",
    fieldName: "AttachmentFiles",
    minWidth: 50,
    onRender(item: ISPListItem): JSX.Element {
      const attachment: attachments[] = item.AttachmentFiles;
      const attachmentLinks: JSX.Element[] = attachment.map((file) => (
        <li>
          <Link href={file.ServerRelativeUrl}>{file.FileName}</Link>
        </li>
      ));
      return <ul>{attachmentLinks}</ul>;
    },
  },
];

export default class Tasks extends React.Component<ITasksProps, ITasksState> {
  protected SPClient: SPClient;
  constructor(props) {
    super(props);
    this.state = { items: [] };
    this.SPClient = new SPClient(this.props.webPartContext);
    this._getItems();
  }
  private _getItems() {
    if (Environment.type == EnvironmentType.Local) {
      MockSPClient.get("http://").then((data: ISPListItem[]) => {
        this.setState({ items: [...data] });
      });
    } else {
      this.SPClient.getRecentItems().then((data: ISPListItem[]) => {
        this.setState({ items: [...data] });
      });
    }
  }

  private _showPanel(item: any, idx: number, e: any) {
    console.log("Panel should Open now for item with Index: ", idx);
  }

  public render(): React.ReactElement<ITasksProps> {
    const items = this.state.items;
    return (
      <div>
        <div>
          <DetailsList
            items={items}
            columns={selectedColumns}
            onActiveItemChanged={this._showPanel.bind(this)}
          />
        </div>
        <div>
          <h4>Add New Assignments</h4>
        </div>
        <NewTask
          webPartContext={this.props.webPartContext}
          onSaveCompleted={this._getItems.bind(this)}
        />
      </div>
    );
  }
}
