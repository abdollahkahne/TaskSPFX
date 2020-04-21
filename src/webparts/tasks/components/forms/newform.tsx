import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import { DatePicker } from "office-ui-fabric-react/lib/DatePicker";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";

import SPClient from "../../SPClient/SPClient";

import { Attachment } from "./Attachment/FilePicker";

import * as React from "react";
import { ISPListItem } from "../../SPClient/ISPListItem";

interface INewTaskProps {
  webPartContext: any;
  onSaveCompleted: any;
}

interface INewTaskState {
  attachements: any;
  title: string;
  assignee: any;
  dueDate: Date;
  note: string;
}

export default class NewTask extends React.Component<
  INewTaskProps,
  INewTaskState
> {
  protected SPClient: SPClient;
  private _peoplePicker: PeoplePicker;
  private _Attachments: Attachment;
  public constructor(props) {
    super(props);
    this.state = {
      title: "",
      dueDate: null,
      assignee: "",
      note: "",
      attachements: [],
    };
    this.SPClient = new SPClient(this.props.webPartContext);
  }

  private _taskTitleChanged(e: any) {
    this.setState({ title: e.target.value });
  }

  private _taskNoteChanged(e: any) {
    this.setState({ note: e.target.value });
  }

  private _saveAttachments(files: any) {
    this.setState({ attachements: files });
  }

  private _assigneeSelected(person) {
    console.log(person);

    this.setState({ assignee: person[0].id });
  }

  private _taskDueDateChanged(date) {
    this.setState({ dueDate: date });
  }

  private _createAssignment() {
    if (!this.state.title || !this.state.assignee) {
      alert("Please Complete Required Fields");
      return;
    }
    console.log(this.state);
    let item: ISPListItem = {
      Title: this.state.title,
      Assignee: { Id: this.state.assignee },
      Note: this.state.note,
      DueDate: this.state.dueDate,
    };
    this.SPClient.createItem(item)
      .then((id) => {
        return this.SPClient.uploadAttachments(this.state.attachements, id);
      })
      .then(() => {
        this.props.onSaveCompleted();
        this._refreshForm();
      });
  }
  private _refreshForm() {
    this.setState({
      attachements: [],
      dueDate: null,
      assignee: "",
      title: "",
      note: "",
    });
    this._peoplePicker.setState({ selectedPersons: [] });
    this._Attachments.setState({ value: "" });
  }

  public render(): React.ReactElement<INewTaskProps> {
    return (
      <div>
        <TextField
          label="Task Name"
          onChange={this._taskTitleChanged.bind(this)}
          required={true}
          value={this.state.title}
        />
        <PeoplePicker
          titleText="Assigned To"
          context={this.props.webPartContext}
          principalTypes={[PrincipalType.User]}
          isRequired={true}
          selectedItems={this._assigneeSelected.bind(this)}
          ensureUser={true}
          ref={(e) => (this._peoplePicker = e)}
        />
        <TextField
          onChange={this._taskNoteChanged.bind(this)}
          label="Note"
          multiline
          value={this.state.note}
        />
        <DatePicker
          label="Due Date"
          onSelectDate={this._taskDueDateChanged.bind(this)}
          value={this.state.dueDate}
        />
        <br />
        <Attachment
          onSave={this._saveAttachments.bind(this)}
          ref={(e) => (this._Attachments = e)}
        />
        <br />
        <PrimaryButton
          text="Save Assignment"
          onClick={this._createAssignment.bind(this)}
        />
      </div>
    );
  }
}
