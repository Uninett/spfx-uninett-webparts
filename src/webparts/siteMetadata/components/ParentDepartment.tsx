import * as React from 'react';
import * as $ from 'jquery';
import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import { IParentDepartmentState } from './IParentDepartmentState';
import { constructor } from "react";
import * as strings from 'SiteMetadataWebPartStrings';
import { Key } from "react";

export interface IParentDepartmentProps {
  context: WebPartContext;
  onChanged: (option: IDropdownOption, index?: number) => void;
  defaultSelectedKey: string;
  orderSiteURL: string;
}


class ParentDepartment extends React.Component<IParentDepartmentProps, IParentDepartmentState> {

  constructor(props: any) {
    super(props);
    this.state = {
      availableParentDepartments: []
    };
  }

  public componentWillMount() {
    this._getChoiceFields();
  }

  public render() {
    return (
    <div>
      {this.state.availableParentDepartments && this.state.availableParentDepartments.length > 0 ? 
      <Dropdown
          label={strings.ParentDepartment}
          options={this.state.availableParentDepartments}
          onChanged={this.props.onChanged}
          defaultSelectedKey={this.props.defaultSelectedKey}
          key={'1'}
        /> : <Dropdown
              label={strings.ParentDepartment}
              options={[]}
              onChanged={this.props.onChanged}
              disabled={true}
              key={'2'}
            />}        
    </div>);
  }

  private ensureRequestDigest = (): Promise<string> => {
    let { serviceScope, pageContext } = this.props.context;
    // pub.selectedLibrary = {
    //   id: decodedQueryString("SPListId"),
    //   listId: pub.cleanGuid(decodedQueryString("SPListId")),
    //   siteUrl: decodedQueryString("SPSiteUrl"),
    //   urlDir: decodedQueryString("SPListUrlDir"),
    //   url: decodedQueryString("SPSiteUrl") + "/" + decodedQueryString("SPListUrlDir"),
    //   contentTypes: {}
    // };
    // monkey patch #__REQUESTDIGEST element
    var __REQUESTDIGEST = document.getElementById("__REQUESTDIGEST");
    if (!__REQUESTDIGEST) {
      __REQUESTDIGEST = document.createElement("input");
      __REQUESTDIGEST.setAttribute("id", "__REQUESTDIGEST");
      __REQUESTDIGEST.setAttribute("name", "__REQUESTDIGEST");
      __REQUESTDIGEST.setAttribute("type", "hidden");
      const digestCache: IDigestCache = serviceScope.consume(
        DigestCache.serviceKey
      );
      return digestCache
        .fetchDigest(pageContext.web.serverRelativeUrl);
      // .then((digest: string): void => {
      //   // use the digest here
      //   __REQUESTDIGEST.setAttribute("value", digest);
      //   document.body.appendChild(__REQUESTDIGEST);

      //   return digest;
      // })
    }
  }

    private _getChoiceFields = (): void => {
      var relativeSiteUrl = this.props.context.pageContext.web.serverRelativeUrl;
      var listName = "Bestillinger";
      var choiceFieldName = "KDTOParentDepartment";
      var choiceFieldId = "5639ffa9-c62e-4513-b7c7-ccca2b5e92c2";
      var restEndPointUrl = this.props.orderSiteURL + "/_api/web/lists/getbytitle('" + listName + "')//fields(guid'" + choiceFieldId + "')/Choices";

      this.ensureRequestDigest().then(requestDigest => {
          $.ajax({
            contentType: 'application/json',
            url: restEndPointUrl,
            type: "GET",
            headers: {
              "Accept": "application/json;odata=nometadata",
              "content-type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest
            },
            success: (data) => {
              var result = data.value;
              var stateValue = result.map( val => ({ key: val, text: val}) );
              this.setState({availableParentDepartments: stateValue});
            },
            error: (err) => {
              
            }
          });
      });
    }
}

export { ParentDepartment };