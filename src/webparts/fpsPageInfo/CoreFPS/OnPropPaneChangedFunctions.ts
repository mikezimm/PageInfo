
import { IMinCustomHelpProps } from "@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup";
import { _LinkIsValid } from '@mikezimm/npmfunctions/dist/Links/AllLinks';

export async function validateDocumentationUrl ( thisProps: IMinCustomHelpProps, propertyPath: string , newValue: any) {

    if ( propertyPath === 'documentationLinkUrl' || propertyPath === 'fpsImportProps' ) {
        thisProps.documentationIsValid = await _LinkIsValid( newValue ) === "" ? true : false;
        console.log( `${ newValue ? newValue : 'Empty' } Docs Link ${ thisProps.documentationIsValid === true ? ' IS ' : ' IS NOT ' } Valid `);

    } else {
        if ( !thisProps.documentationIsValid ) { thisProps.documentationIsValid = false; }

    }

}