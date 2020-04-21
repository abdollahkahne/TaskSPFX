import * as React from "react";
interface IAttachmentProps {
  onSave?: any;
}
interface IAttachmentState {
  value: any;
}
export class Attachment extends React.Component<
  IAttachmentProps,
  IAttachmentState
> {
  public constructor(props) {
    super(props);
    this.state = { value: "" };
  }
  private _onChange(e: any) {
    this.setState({ value: e.target.value });
    this.props.onSave(e.target.files);
  }
  public render(): React.ReactElement {
    return (
      <div>
        <label>Add Attachments: </label>
        <br />
        <input
          type="file"
          value={this.state.value}
          onChange={this._onChange.bind(this)}
          multiple
        />
      </div>
    );
  }
}
