import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

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
        defPinState: "pinFull",
        forcePinState: true,
        selectedProperties: [
            "Title"
        ],
        h1Style: "background:#e3e3e3;color:#005495;padding:10px 20px",
        pageInfoStyle: '\"paddingBottom\":\"20px\",\"backgroundColor\":\"#dcdcdc\";\"borderLeft\":\"solid 3px #c4c4c4\"',

        feedbackEmail: `ae57524a.${window.location.hostname}.onmicrosoft.com@amer.teams.ms`,
        panelMessageDescription1: 'Finance Manual Help and Contact',
        panelMessageSupport: 'Contact RE for Finance Manual content',
        panelMessageDocumentation: 'Contact MZ for Web part questions',
        panelMessageIfYouStill: '',
        documentationLinkDesc: 'Finance Manual Help site',
        documentationLinkUrl: '/sites/FinanceManual/Help',
        documentationIsValid: true,
        supportContacts: [ FinancialManualContacts ],

        bannerStyleChoice: "lock",
        bannerStyle: '{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":\"larger\",\"fontWeight\":600,\"fontStyle\":\"normal\",\"padding\":\"0px 10px\",\"height\":\"48px\",\"cursor\":\"pointer\"}',
        bannerCmdStyle: '{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":16,\"fontWeight\":\"normal\",\"fontStyle\":\"normal\",\"padding\":\"7px 4px\",\"marginRight\":\"0px\",\"borderRadius\":\"5px\",\"cursor\":\"pointer\"}',

    }
};

export const PreConfiguredPrpos : IAllPreConfigSettings = {
    forced: [ ForceFinancialManual ],
    preset: [],
};