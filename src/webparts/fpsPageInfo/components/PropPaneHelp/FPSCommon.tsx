import * as React from 'react';
import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';

import styles from './PropPanelHelp.module.scss';

import { Pivot, PivotItem, IPivotItemProps, PivotLinkFormat, PivotLinkSize,} from 'office-ui-fabric-react/lib/Pivot';

import { defaultBannerCommandStyles, } from "@mikezimm/npmfunctions/dist/HelpPanel/onNpm/defaults";

const CSSOverRideWarning = <div style={{fontSize: 'large' }}>
  <div className={ styles.topic} style={{fontSize: 'large' }}><mark>NOTICE</mark></div>
  <div>Any 3rd party app that modifies the page styling (like these) are using undocumented tricks </div>
  <div>  - - <b>WHICH ARE SUBJECT TO BREAK without notice by Microsoft</b>.</div>
  <div>These settings are applied after this web part loads.</div>
  <div><b>Users may briefly see the original styling</b>.  Especially if their connection is slow or your page takes a long time to load.</div>
</div>;

const DeveloperWarning = <div style={{fontSize: 'large' }}>
  <div className={ styles.topic} style={{fontSize: 'large' }}><mark>NOTICE</mark></div>
  <div>ONLY turn these on IF you know what you are doing and need them.</div>
</div>;

export const VisitorHelp = <PivotItem headerText={ 'Visitor Help' } > 
    <div className={ styles.helpContent}>
    <div className={ styles.topic}>Full Help Panel Audience</div>
    <div>This gives you control who can see the entire <b>More Information</b> panel in the Help Banner bar.</div>
    <div>People who have less rights than this will only see the content you add via the property pane.</div>

    <div className={ styles.topic}>Panel Description</div>
    <div>Personalized heading message you give you your users.</div>

    <div className={ styles.topic}>Support Message</div>
    <div>Optional message to give users for support.</div>

    <div className={ styles.topic}>Documentation message</div>
    <div>Message you can give users directly above the documentation link</div>

    <div className={ styles.topic}>Paste a Documentation link</div>
    <div>We require a valid SharePoint link where you store further information on using this web part.</div>

    <div className={ styles.topic}>Documentation Description</div>
    <div>Optional text that the user sees as the Documentation Link text</div>

    <div className={ styles.topic}>Support Contacts</div>
    <div>Use of this web part requires a current user to be identified for support issues.</div>

    </div>
</PivotItem>;


export const BannerHelp = <PivotItem headerText={ 'Banner' } > 
    <div className={ styles.helpContent}>
    <div className={ styles.topic} style={{ textDecoration: 'underline' }}>FPS Banner - Basics</div>
    <div className={ styles.topic}>Show Banner</div>
    <div>May allow you to hide the banner.  If toggle disabled, it is required.</div>

    <div className={ styles.topic}>Optional Web Part Title</div>
    <div>Add Title text to the web part banner.</div>
    <div>Depending on the web part, this may not be editable.</div>

    <div className={ styles.topic}>More Info text-button</div>
    <div>Customize the More Information text/Icon in the right of the banner.</div>

    <div className={ styles.topic} style={{ textDecoration: 'underline' }}>FPS Banner - Navigation</div>
    <div className={ styles.topic}>Show 'Go to Home Page' <Icon iconName='Home'></Icon> Icon</div>
    <div>Displays the <Icon iconName='Home' style={ defaultBannerCommandStyles }></Icon> when you are not on the site's home page.</div>

    <div className={ styles.topic}>Show 'Go to Parent Site' <Icon iconName='Up'></Icon> Icon</div>
    <div>Displays the <Icon iconName='Up' style={ defaultBannerCommandStyles }></Icon> when you are not on the site's home page.</div>

    <div className={ styles.topic}>Gear, Go to Home, Parent audience</div>
    <div>Minimum permissions requied to see the Home and Parent site icons.</div>
    <div>Use this to hide buttons from visitors if your ALV Financial Manual Web part is more of a single page app and you want to hide the site from a typical visitor.</div>
    <div>NOTE:  Site Admins will always see the icons.</div>
    <ul>
        <li>Site Owners: have manageWeb permissions</li>
        <li>Page Editors: have addAndCustomizePages permissions</li>
        <li>Item Editors: have addListItems permissions</li>
    </ul>
    
    
    

    <div className={ styles.topic} style={{ textDecoration: 'underline' }}>Theme options</div>
    <div><mark><b>NOTE:</b></mark> May be required depending on our policy for this web part</div>
    <div>Use dropdown to change your theme for the banner (color, buttons, text)</div>

    <div className={ styles.topic}>Banner Hover Effect</div>
    <div>Turns on or off the Mouse Hover effect.  If Toggle is off, the banner does not 'Fade In'.  Turn off if you want a solid color banner all the time.</div>

    </div>
</PivotItem>;


export const FPSBasicHelp = <PivotItem headerText={ 'FPS Basic' } > 
    <div className={ styles.helpContent}>

    { CSSOverRideWarning }

    <div className={ styles.topic}>Hide Quick Launch</div>
    <div>As of April 2022, MSFT allows you to modify quick launch in Site Gear 'Change the look'</div>
    <div>Only use this option if you want the Quick launch on the site as a whole but not on the page this web part is installed on.</div>
    
    <div className={ styles.topic}>All Sections <b>Max Width</b> Toggle and slider</div>
    <div>Over-rides out of the box max width on page sections.</div>

    <div className={ styles.topic}>All Sections <b>Margin</b> Toggle and slider</div>
    <div>Over-rides out of the box top and bottom section margin.</div>

    <div className={ styles.topic}>Hide Toolbar - while viewing</div>
    <div>Hidden:  Will hide the page toolbar (Edit button) when loading the page.</div>
    <div><b>Only use if you know what you are doing :)</b></div>
    <div><mark>WARNING</mark>.  <b>Add ?tool=true to the Url</b> and reload the page in order to edit the page.  You <b>CAN NOT SEE THESE INSTRUCTIONS</b> unless you add ?tool=true to the page</div>

    </div>
</PivotItem>;

export const FPSExpandHelp = <PivotItem headerText={ 'FPS Expand' } > 
    <div className={ styles.helpContent}>

    { CSSOverRideWarning }

    <div className={ styles.topic}><b></b>Enable Expandoramic Mode</div>
    <div><b></b>Enables the Expandoramic toggle (diagonal arrow icon in upper left of Header.</div>

    <div className={ styles.topic}><b></b>Page load default</div>
    <div><b></b>Determines the format when loading the page.</div>
    <ul>
        <li>Normal:  Webpart DOES NOT AUTO expand when loading the page</li>
        <li>Expanded:  Page loads with webpart expanded</li>
        <li>Whenever you 'Edit' the page, you may need to manually shrink webpart to see the page and webpart properties.</li>
    </ul>

    <div className={ styles.topic}><b></b>Expandoramic Audience</div>
    <div><b>NOTE:</b> Site Admins will always see all icons regardless of the Toggles or the audience.</div>
    <ul>
        <li>Site Owners: have manageWeb permissions</li>
        <li>Page Editors: have addAndCustomizePages permissions</li>
        <li>Item Editors: have addListItems permissions</li>
    </ul>

    <div className={ styles.topic}><b>Style options and Hover Effect</b> are for SharePoint IT use only.</div>
    <div><b></b></div>

    <div className={ styles.topic}>Padding</div>
    <div>Adjusts the padding around the webpart.  20px minimum.</div>

    </div>
</PivotItem>;

export const SinglePageAppHelp = <PivotItem headerText={ 'Single Page Apps' } > 
    <div className={ styles.helpContent}>

    <div className={ styles.topic}>Before you start!</div>

    <div className={ styles.topic}>If you plan to build a full page app (Full expand web part at load time)</div>
    <div>
        Be sure to follow these steps to improve performance and minimize any styling issues and delays:
        <ol>
        <li>Create a page from 'Apps' Template when you first create a page</li>
        <ul>
            <li>This will remove all navigation from the page, make the web part full page and load faster.</li>
        </ul>

        <li>IF NOT, then Start with a <b>Communication Site</b></li>
        <ul>
            <li>This is the only site that allows true 'Full Width webparts'</li>
        </ul>
        <li><b>Clear the home page completely</b> (do not have any other webparts)</li>
        <li>Minimize what SharePoint loads
        <ol style={{ listStyleType: 'lower-alpha' }}>
            <li>Go to Gear</li>
            <li>Click 'Change the look'</li>
            <li>Click Header
            <ul>
            <li>Set Layout to minimal</li>
            <li>Set 'Site title visiblity' to off</li>
            <li>Remove your site logo</li>
            <li>Save Header settings</li>
            </ul></li>
            <li>Click Navigation
            <ul>
            <li>Turn off Site Navigation</li>
            </ul></li>
        </ol></li>
        <li>Add SecureScript in the first Full Width section</li>
        </ol>
    </div>
    </div>
</PivotItem>;

export const ImportHelp = <PivotItem headerText={ 'Import' } > 
    <div className={ styles.helpContent}>
        <div className={ styles.topic}>If Available in this web part...</div>
        <div>It allows you to paste in values from the same webpart from a different page.</div>
        <div>To Export web part settings</div>
        <ol>
        <li>Click on 'More Information' in the Web Part Banner</li>
        <li>Click the Export tab <Icon iconName='Export' style={ defaultBannerCommandStyles }></Icon> (last tab in the Help Panel)</li>
        <li>Hover over Export Properties row</li>
        <li>Click the blue paper/arrow icon on the right side of the row to 'Export' the properties</li>
        <li>Edit this page and web part</li>
        <li>Paste properties into the Import properties box</li>
        </ol>

    </div>
</PivotItem>;
