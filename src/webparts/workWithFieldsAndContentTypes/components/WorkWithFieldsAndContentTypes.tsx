import * as React from "react";
import { IWorkWithFieldsAndContentTypesProps } from "./IWorkWithFieldsAndContentTypesProps";
import { sp } from "@pnp/sp/presets/all";
import { PrimaryButton } from "office-ui-fabric-react";

export default class WorkWithFieldsAndContentTypes extends React.Component<
  IWorkWithFieldsAndContentTypesProps,
  {}
> {
  public async componentDidMount() {
    sp.setup(this.props.context);
  }

  private addTypedItem = async () => {
    // Determine the target content type ID (assuming it does exist)
    const contentTypes = await sp.web.lists
      .getByTitle(this.props.listTitle)
      .contentTypes.get();
    const targetContentType = contentTypes.filter(
      (v) => v.Name === "PiaSysCustomer"
    )[0];

    // Determine the ID of the reference User
    const referenceUser = await sp.web.ensureUser(
      "junaid@sfnwq.onmicrosoft.com"
    );
    const referenceUserId = referenceUser.data.Id;

    const iar = await sp.web.lists.getByTitle(this.props.listTitle).items.add({
      ContentTypeId: targetContentType.Id.StringValue,
      Title: "Sample Customer 2",
      RefererId: referenceUserId,
      Address: "ABC street",
    });

    console.log("iar", iar);
  };

  public render(): React.ReactElement<IWorkWithFieldsAndContentTypesProps> {
    // const {
    //   listTitle,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <div>
        <span>Welcome to Sharepoint!</span>
        <p>Customize SharePoint experiences using web parts.</p>
        <PrimaryButton text="Add a Typed item" onClick={this.addTypedItem} />
        &nbsp;
      </div>
      // <section className={`${styles.workWithFieldsAndContentTypes} ${hasTeamsContext ? styles.teams : ''}`}>
      //   <div className={styles.welcome}>
      //     <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
      //     <h2>Well done, {escape(userDisplayName)}!</h2>
      //     <div>{environmentMessage}</div>
      //     <div>Web part property value: <strong>{escape(description)}</strong></div>
      //   </div>
      //   <div>
      //     <h3>Welcome to SharePoint Framework!</h3>
      //     <p>
      //       The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
      //     </p>
      //     <h4>Learn more about SPFx development:</h4>
      //     <ul className={styles.links}>
      //       <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
      //       <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
      //       <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
      //       <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
      //       <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
      //       <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
      //       <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
      //     </ul>
      //   </div>
      // </section>
    );
  }
}
