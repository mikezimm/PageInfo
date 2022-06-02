import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { IFpsPageInfoWebPartProps } from "./IFpsPageInfoWebPartProps";

import { IPreConfigSettings, IAllPreConfigSettings } from '@mikezimm/npmfunctions/dist/PropPaneHelp/PreConfigFunctions';
import { encrptMeOriginalTest } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/logTest';
import { ContALVFMContent, ContALVFMWebP } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/constants';

const FinancialManualContacts: IPropertyFieldGroupOrPerson = {
    id: '1',
    description: '',
    fullName: 'Financial Manual Support team',
    login: '',
    email: `ae57524a.${window.location.hostname}.onmicrosoft.com@amer.teams.ms`,
    // jobTitle?: string;
    // initials?: string;
    imageUrl: null,
};

export const ForceFinancialManual : IPreConfigSettings = {
    location: '/sites/financemanual/',
    props: {
        // Pin Me props that are not preset in manifest.json
        defPinState: "pinFull",
        forcePinState: true,

        // Web part styling props that are not preset in manifest.json
        h1Style: "background:#e3e3e3;color:#005495;padding:10px 20px",
        pageInfoStyle: '\"paddingBottom\":\"20px\",\"backgroundColor\":\"#dcdcdc\";\"borderLeft\":\"solid 3px #c4c4c4\"',

        // Properties props that are not preset in manifest.json
        selectedProperties: [
            "ALGroup",
            "DocumentType",
            "Functions",
            "Processes",
            "ReportingSections",
            "StandardDocuments",
            "Topics",
        ],

        // Visitor Panel props that are not preset in manifest.json
        fullPanelAudience: 'Page Editors',
        panelMessageDescription1: 'Finance Manual Help and Contact',
        panelMessageSupport: `Contact ${encrptMeOriginalTest( ContALVFMContent )} for Finance Manual content`,
        panelMessageDocumentation: `Contact ${encrptMeOriginalTest( ContALVFMWebP )}  for Web part questions`,
        panelMessageIfYouStill: '',
        documentationLinkDesc: 'Finance Manual Help site',
        documentationLinkUrl: '/sites/FinanceManual/Help',
        documentationIsValid: true,
        supportContacts: [ FinancialManualContacts ],

        // FPS Banner Basics
        bannerTitle: 'Page Info',
        infoElementChoice: "IconName=Unknown",
        infoElementText: "Question mark circle",
        feedbackEmail: `ae57524a.${window.location.hostname}.onmicrosoft.com@amer.teams.ms`,

        // FPS Banner Navigation
        showGoToHome: true,
        showGoToParent: true,

        // Banner Theme props that are not preset in manifest.json
        bannerStyleChoice: 'corpDark1',
        bannerStyle: '{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":\"larger\",\"fontWeight\":600,\"fontStyle\":\"normal\",\"padding\":\"0px 10px\",\"height\":\"48px\",\"cursor\":\"pointer\"}',
        bannerCmdStyle: '{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":16,\"fontWeight\":\"normal\",\"fontStyle\":\"normal\",\"padding\":\"7px 4px\",\"marginRight\":\"0px\",\"borderRadius\":\"5px\",\"cursor\":\"pointer\"}',
        lockStyles: true,

        relatedStyle: '',

        related1description: 'Standards',
        related1showItems: true,
        related1isExpanded: true,
        related1web: '/sites/financemanual/manual',
        related1listTitle: 'Site Pages',
        related1restFilter: 'StandardDocumentsId eq {{PageId}}',
        related1linkProp: 'File/ServerRelativeUrl', // aka FileLeaf to open file name, if empty, will just show the value
        related1displayProp: 'Title',

        related2description: 'Supporting Documents',
        related2showItems: true,
        related2isExpanded: false,
        related2web: '/sites/financemanual/manual',
        related2listTitle: 'SupportDocuments',
        related2restFilter: 'StandardDocumentsId eq {{PageId}}',
        related2linkProp: 'File/ServerRelativeUrl', // aka FileLeaf to open file name, if empty, will just show the value
        related2displayProp: 'FileLeafRef',

        pageLinksdescription: 'Images and Links',
        pageLinksshowItems: true,
        pageLinksisExpanded: false,
        canvasLinks: true,
        canvasImgs: true,
        pageLinksweb: '/sites/financemanual/manual',
        pageLinkslistTitle: 'Site Pages',
        pageLinksrestFilter: 'ID eq {{PageId}}',
        pageLinkslinkProp: 'File/ServerRelativeUrl', // aka FileLeaf to open file name, if empty, will just show the value
        pageLinksdisplayProp: 'FileLeafRef',
        
    }
};


export const PresetFinancialManual : IPreConfigSettings = {
    location: '/sites/financemanual/',
    props: {
        homeParentGearAudience: 'Everyone',
    }
};

export const PresetSomeRandomSite : IPreConfigSettings = {
    location: '/sites/SomeRandomSite/',
    props: {
        homeParentGearAudience: 'Some Test Value',
    }
};

export const PreConfiguredProps : IAllPreConfigSettings = {
    forced: [ ForceFinancialManual ],
    preset: [ PresetFinancialManual, PresetSomeRandomSite ],
};
