import * as React from 'react';
import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';

import { Pivot, PivotItem, IPivotItemProps, PivotLinkFormat, PivotLinkSize,} from 'office-ui-fabric-react/lib/Pivot';

import { IQuickCommands } from '@mikezimm/npmfunctions/dist/QuickCommands/IQuickCommands';

import { IRefinerRulesStrs, IRefinerRulesInts, IRefinerRulesNums, IRefinerRulesTime, IRefinerRulesUser, IRefinerRulesEXPE, IRefinerRulesNone } from '@mikezimm/npmfunctions/dist/Refiners/IRefiners';
import { RefinerRulesStrs, RefinerRulesInts, RefinerRulesNums, RefinerRulesTime, RefinerRulesUser, RefinerRulesEXPE, RefinerRulesNone } from '@mikezimm/npmfunctions/dist/Refiners/IRefiners';

import { gitRepoALVFinManSmall } from '@mikezimm/npmfunctions/dist/Links/LinksRepos';

import { defaultBannerCommandStyles, } from "@mikezimm/npmfunctions/dist/HelpPanel/onNpm/defaults";

import ReactJson from "react-json-view";
import { FontWeights } from 'office-ui-fabric-react';
import { BannerHelp, FPSBasicHelp, FPSExpandHelp, ImportHelp, SinglePageAppHelp, VisitorHelp } from '@mikezimm/npmfunctions/dist/PropPaneHelp/FPSCommonOnNpm';
import { FPSBorderClasses, FPSHeadingNumberClasses, FPSEmojiClasses, FPSMiscClasses, FPSHeadingClasses } from '../HeadingCSS/FPSTagFunctions';

require('@mikezimm/npmfunctions/dist/PropPaneHelp/FPSCommonOnNpm/PropPanelHelp.css');

const SampleViewJSON : any = [
  {
    "name": "Author/Title",
    "displayName": "Created by",
    "minWidth": 50
  },
  {
    "name": "FileRef",
    "displayName": "FileLeafRef",
    "maxWidth": 50,
    "linkPropertyName": "goToItemPreview"
  },
];

// const SampleCommands: any = {
//   "label": "Press Me",
//   "primary": false,
//   "confirm": "Are you sure?",
//   "alert": "We made our updates!",
//   "console": "Message to browser console",
//   "panelMessage": "Complete Panel Text",
//   "icon": "User",
//   "updateItem": {
//     "DueDate": "[today+14]",
//     "AssignedToId": "[Me]",
//     "Status": "In Process",
//     "ReviewDays": 99,
//     "Body": "Hi! It's [Today+3] and I'm $MyName$"
//   },
//   "showWhenEvalTrue": "item.AssignedToId !== sourceUserInfo.Id"
// };

const SampleCommands: any = {
  "buttons": [[{
      "label": "Press Me",
      "primary": false,
      "confirm": "Are you sure?",
      "alert": "We made our updates!",
      "console": "Message to browser console",
      "panelMessage": "Complete Panel Text",
      "icon": "User",
      "updateItem": {
        "DueDate": "[today+14]",
        "AssignedToId": "[Me]",
        "Status": "In Process",
        "ReviewDays": 99,
        "Body": "Hi! It's [Today+3] and I'm $MyName$"
      },
      "showWhenEvalTrue": "item.AssignedToId !== sourceUserInfo.Id"
    }
  ]],
  "fields": [],

};
 
const SampleCharts: any = [
  {
    "primaryField": "Id",
    "title": "Count of items",
    "stat": "count",
    "chartTypes": [
      "pareto-dec",
      "stacked-column-labels"
    ]
  }
];

const UserColumnRestPropertiesSPO : string[] = [ 'Title', 'Name', 'EMail', 'FirstName', 'UserName', 'ID', 'SipAddress', 'Office', 'Modified', 'Created', ];
const UserColumnRestPropertiesSPONOTWORK : string[] = [ 'MobilePhone', 'Department', 'JobTitle', 'WorkPhone', 'ImnName', 'NameWithPicture', 'NameWithPictureAndDetails', 'ContentTypeDisp', ];

export function putObjectIntoJSON ( obj: any, name: string = null ) {
  // return <ReactJson src={ obj } name={ 'panelItem' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>;
  return <ReactJson src={ obj } name={ name } collapsed={ false } displayDataTypes={ false } displayObjectSize={ false } enableClipboard={ true } style={{ padding: '20px 0px' }} theme= { 'rjv-default' } indentWidth={ 2}/>;
}

const PleaseSeeWiki = <p>Please see the { gitRepoALVFinManSmall.wiki }  for more information</p>;

const tenantServiceRequestURL = `https://servicenow.${window.location.hostname}.com/`;
const RequestStorageHere = <span>Please request storage <a href={tenantServiceRequestURL} target="_blank">here in Service Now.</a></span>;

const LinkFindInternalName = <a href="https://tomriha.com/what-is-sharepoint-column-internal-name-and-where-to-find-it/" target="_blank">Finding Internal Name of a column</a>;

const ShowCodeIcon = <Icon iconName={ 'Code' } title='ShowCode icon' style={ defaultBannerCommandStyles }></Icon>;
const CheckReferences = <Icon iconName={ 'PlugDisconnected' } title='Check Files' style={ defaultBannerCommandStyles }></Icon>;
const ShowRawHTML = <Icon iconName={ 'FileCode' } title='Show Raw HTML here' style={ defaultBannerCommandStyles }></Icon>;

const padRight15: React.CSSProperties = { paddingRight: '15px' };
const padRight40: React.CSSProperties = { paddingRight: '40px' };

export const WebPartHelpElement = <div>
  <Pivot 
          linkFormat={PivotLinkFormat.links}
          linkSize={PivotLinkSize.normal}
      //   style={{ flexGrow: 1, paddingLeft: '10px' }}
      //   styles={ null }
      //   linkSize= { pivotOptionsGroup.getPivSize('normal') }
      //   linkFormat= { pivotOptionsGroup.getPivFormat('links') }
      //   onLinkClick= { null }  //{this.specialClick.bind(this)}
      //   selectedKey={ null }
      >
      <PivotItem headerText={ 'Pin Me' } > 
        <div className={ 'fps-pph-content' }>
          <div className={ 'fps-pph-topic' }>Default Location</div>
          <div>
            <li><b>normal - </b>Web part loads on page where you put it</li>
            <li><b>Pin Expanded - </b>Web part loads Pinned in upper right corner fully expanded</li>
            <li><b>Pin Collapsed - </b>Web part loads Pinned in upper right corner collapsed</li>
          </div>
          <div className={ 'fps-pph-topic' }>Force Pin State</div>
          <div>
            <li><b>Let user change - </b>End user can move the web part from Pinned to Normal location at any time</li>
            <li><b>Enforce no Toggle - </b>End user can not toggle the position of the web part.
              <p>With Enforcing pin, the end user will always be able to expand or collapse the web part.</p>
              <p>Be sure to test experience by loading the page with the browser shrunk to size of a phone.  Consider end user experience trying to navigate your page.</p>
            </li>
          </div>       
        </div>
      </PivotItem>
    
      <PivotItem headerText={ 'Table of Contents' } > 
        <div className={ 'fps-pph-content' }>

          <div className={ 'fps-pph-topic' }>Show TOC - Table of Contents</div>
          <div>Shows Table of Contents component which builds Header navigation links.</div>

          <div className={ 'fps-pph-topic' }>Default state</div>
          <div>How the web part initially loads.</div>

          <div className={ 'fps-pph-topic' }>TOC Heading or Title</div>
          <div><b>Recommended - </b>Header text above TOC.  If you have text here, you can expand and collapse this section of the web part</div>

          <div className={ 'fps-pph-topic' }>Min heading to show</div>
          <div>Select minimum heading levels to show in TOC.  h1 will only show Heading1.  h2 will show Heading1 and Heading2.  h3 will show Heading1, Heading2 and Heading3</div>
        </div>
      </PivotItem>
    
      <PivotItem headerText={ 'Properties' } > 
        <div className={ 'fps-pph-content' }>

          <div className={ 'fps-pph-topic' }>Show Created-Modified Properties - from the page</div>
          <div>Shows out of the box page properties.</div>

          <div className={ 'fps-pph-topic' }>Show Approval Status Properties - from the page</div>
          <div>Shows page approval status information.</div>

          <div className={ 'fps-pph-topic' }>Show Custom Properties - from the page</div>
          <div>Shows columns-properties on this site page.</div>
          <div>Use the +Add and -Delete buttons to add or delete page properties you want to show.</div>

          <div className={ 'fps-pph-topic' }>TOC Heading or Title</div>
          <div><b>Recommended - </b>Header text above Properties.  If you have text here, you can expand and collapse this section of the web part</div>

          <div className={ 'fps-pph-topic' }>Default state</div>
          <div>How the web part initially loads.</div>
        </div>
      </PivotItem>

      <PivotItem headerText={ 'Web part styles' } > 
        <div className={ 'fps-pph-content' }>

          <div className={ 'fps-pph-topic' }>Heading 1, Heading 2, Heading 3, Styles</div>
          <div>Apply classes and styles to respective Heading elements on the page.   You can combine both classes and styles as shown below</div>
          <div>.dottedTopBotBorder;color:'red' %lt;== this will add dotted top and bottom border class and add font-color: red style to the heading.</div>
          
          <div style={{ display: 'flex' }}>
                <div style={ padRight40 }><div className={ 'fps-pph-topic' }>Border Classes</div><ul>
                  { FPSBorderClasses.map( rule => <li>{ '.' + rule }</li> ) }
                </ul></div>
                <div style={ padRight40 }><div className={ 'fps-pph-topic' }>Heading Numb Classes</div><ul>
                  { FPSHeadingNumberClasses.map( rule => <li>{ '.' + rule }</li> ) }
                </ul></div>

                <div style={ padRight40 }><div className={ 'fps-pph-topic' }>Emoji Classes</div><ul>
                  { FPSEmojiClasses.map( rule => <li>{ '.' + rule }</li> ) }
                </ul></div>

                <div style={ padRight40 }><div className={ 'fps-pph-topic' }>Misc Classes</div><ul>
                  { FPSMiscClasses.map( rule => <li>{ '.' + rule }</li> ) }
                </ul></div>
            </div>

          <div className={ 'fps-pph-topic' }>Page Info Style options</div>
          <div>Applies to the container below the banner that contains both the TOC and Props components</div>
          <div>React.CSSProperties string like (with quotes): "fontSize":"larger","color":"red"</div>

          <div className={ 'fps-pph-topic' }>Table of Contents Style options</div>
          <div>Applies to the Table of Contents container</div>
          <div>React.CSSProperties string like (with quotes): "fontWeight":600,"color":"yellow"</div>

          <div className={ 'fps-pph-topic' }>Properties Style options</div>
          <div>Applies to the Properties container</div>
          <div>React.CSSProperties string like (with quotes): "fontWeight":600,"color":"yellow"</div>
        </div>
      </PivotItem>

      { VisitorHelp }
      { BannerHelp }
      { FPSBasicHelp }
      { FPSExpandHelp }
      { SinglePageAppHelp }
      { ImportHelp }

  </Pivot>
</div>;