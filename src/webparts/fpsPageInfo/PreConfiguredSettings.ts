import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { IFpsPageInfoWebPartProps } from "./IFpsPageInfoWebPartProps";

export interface IPreConfigSettings {
    location: string;
    props: any;
}

export interface IAllPreConfigSettings {
    forced: IPreConfigSettings[];
    preset: IPreConfigSettings[];
}

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
        panelMessageSupport: 'Contact RE for Finance Manual content',
        panelMessageDocumentation: 'Contact MZ for Web part questions',
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

    }
};


export const PresetFinancialManual : IPreConfigSettings = {
    location: '/sites/financemanual/',
    props: {

        homeParentGearAudience: 'Everyone',
    }
};

export const PreConfiguredPrpos : IAllPreConfigSettings = {
    forced: [ ForceFinancialManual ],
    preset: [ PresetFinancialManual ],
};

export const ThisSitesPreConfiguredPrpos : IAllPreConfigSettings = {
    forced: [ ForceFinancialManual ],
    preset: [ PresetFinancialManual ],
};

export interface IConfigurationProp {
    location: string;
    prop: string;
    value: any;
    type: 'preset' | 'forced' | 'unk';
    status: 'tbd' |  'valid' | 'preset' | 'force-preset' | 'force-changed' | 'error' | 'unk';

}

export interface ISitePreConfigProps {
    presets: IConfigurationProp[];
    forces: IConfigurationProp[];
}

export function getThisSitesPreConfigProps( thisProps: IFpsPageInfoWebPartProps , serverRelativeUrl: string ) : ISitePreConfigProps {

    let presets: IConfigurationProp[] = [];
    let forces: IConfigurationProp[] = [];

    PreConfiguredPrpos.preset.map( preconfig => {
      if ( serverRelativeUrl.toLowerCase().indexOf( preconfig.location ) > -1 ) {
        Object.keys( preconfig.props ).map( prop => {
          if ( !thisProps[prop] ) { 
            presets.push( { location: preconfig.location, type: 'preset', prop: prop, value: preconfig.props[ prop ], status: 'tbd' });
          }
        });
      }
    });

    PreConfiguredPrpos.forced.map( preconfig => {
      if ( serverRelativeUrl.toLowerCase().indexOf( preconfig.location ) > -1 ) {
        Object.keys( preconfig.props ).map( prop => {
          if ( thisProps[prop] !== preconfig.props[ prop ] ) {
            forces.push( { location: preconfig.location, type: 'forced', prop: prop, value: preconfig.props[ prop ], status: 'tbd' });
          }
        });
      }
    });

    return { presets: presets, forces: forces };

}