import * as React from 'react';
import styles from './ReactCustomPropertyPane.module.scss';
import { IReactCustomPropertyPaneProps } from './IReactCustomPropertyPaneProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import * as pnp from 'sp-pnp-js';
import IReactCustomPropertyPaneState from './IReactCustomPropertyPaneState';
import IPropsItem from './IPropsItem';

import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

export default class ReactCustomPropertyPane extends React.Component<IReactCustomPropertyPaneProps, IReactCustomPropertyPaneState> {

  constructor(props: IReactCustomPropertyPaneProps, state: IReactCustomPropertyPaneState) {
    super(props);
    this.state = {
      listDetails: { Id: '', Url: '' },
      libraryDetails: { Id: '', Url: '' },
      isListPanel: false,
      isLibraryPanel :false
    }

    this._showListDetails = this._showListDetails.bind(this);
    this._showLibraryDetails = this._showLibraryDetails.bind(this);
  }


  private _showListDetails() {
    pnp.sp.web.lists.getByTitle(this.props.listName).views.getByTitle(this.props.listViewName).get().then(v => {
      this.setState({
        listDetails: { Id: v.Id, Url: v.ServerRelativeUrl },
        isListPanel: true,
      });
    })

  }

  private _showLibraryDetails() {
    pnp.sp.web.lists.getByTitle(this.props.libraryName).views.getByTitle(this.props.libraryViewName).get().then(v => {
      console.log("_showLibraryDetails");
      console.log(v.Id);
      console.log(v.ServerRelativeUrl);
      this.setState({
        libraryDetails: { Id: v.Id, Url: v.ServerRelativeUrl },
        isLibraryPanel: true,
      });
    })
  }

  public render(): React.ReactElement<IReactCustomPropertyPaneProps> {
    let { isListPanel,isLibraryPanel, listDetails, libraryDetails } = this.state;

    return (
      <div className={styles.reactCustomPropertyPane}>
        <div className={styles.container}>
          <div className="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">{this.props.title}</span>
              <p className="ms-font-l ms-fontColor-white">{this.props.description}</p>
              <table width="100%">
                <tbody>
                  <tr>
                    <td><DefaultButton description='Opens the List Details Panel' onClick={() => this._showListDetails()}
                      text='List Details' /> </td>
                    <td><DefaultButton description='Opens the Library Details Panel' onClick={() => this._showLibraryDetails()}
                      text='Library Details' /></td>
                  </tr>
                  <tr>
                    <td><Panel
                      isOpen={this.state.isListPanel}
                      onDismiss={() => this.setState({ isListPanel: false })}
                      type={PanelType.large} 
                      headerText='List Details'
                    ><div>
                        <table width="100%">
                          <tbody>
                            <tr>
                              <td>List Name : </td>
                              <td>{this.props.listName}</td>
                            </tr>
                            <tr>
                              <td>View Name : </td>
                              <td>{this.props.listViewName}</td>
                            </tr>
                            <tr>
                              <td>List ID : </td>
                              <td>{this.state.listDetails.Id}  </td>
                            </tr>
                            <tr>
                              <td>List Url </td>
                              <td>{this.state.listDetails.Url}</td>
                            </tr>

                          </tbody>
                        </table>
                      </div>
                    </Panel> </td>
                    <td><Panel
                      isOpen={this.state.isLibraryPanel}
                      onDismiss={() => this.setState({ isLibraryPanel: false })}
                      type={PanelType.medium}
                      headerText='Library Details'
                    ><div>
                        <table width="100%">
                          <tbody>
                            <tr>
                              <td>Library Name : </td>
                              <td>{this.props.libraryName}</td>
                            </tr>
                            <tr>
                              <td>View Name : </td>
                              <td>{this.props.libraryViewName}</td>
                            </tr>
                            <tr>
                              <td>Library ID : </td>
                              <td>{this.state.libraryDetails.Id}  </td>
                            </tr>
                            <tr>
                              <td>Library Url </td>
                              <td>{this.state.libraryDetails.Url}</td>
                            </tr>

                          </tbody>
                        </table>
                      </div>
                    </Panel></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
