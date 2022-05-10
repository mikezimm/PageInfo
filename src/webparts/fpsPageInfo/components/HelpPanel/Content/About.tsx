import * as React from 'react';

//import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../Component/ISinglePageProps';
import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../banner/SinglePage/ISinglePageProps';

import * as devLinks from '@mikezimm/npmfunctions/dist/Links/LinksDevDocs';

import { IRepoLinks } from '@mikezimm/npmfunctions/dist/Links/CreateLinks';

import { convertIssuesMarkdownStringToSpan } from '@mikezimm/npmfunctions/dist/Elements/Markdown';

export const panelVersionNumber = '2022-05-12 -  1.0.0.05'; //Added to show in panel

export function aboutTable( repoLinks: IRepoLinks, showRepoLinks: boolean ) {

    let table : IHelpTable  = {
        heading: 'Version History',
        headers: ['Date','Version','Focus'],
        rows: [],
    };

    /**
     * Security update log 
     * 
     * converting all links and cdns to lower case so casing does miss a flag
     * standardizing all cdn links to start with /sites/ if on tenant
     * standardinzing all tag lings to start with /sites/ if on tenant
     * removing any extra // from both cdns and file links so you cant add extra slash in a url and slip by
     * 
     * Does NOT find files without extensions (like images and also script files.)
     * 
     * WARNING:  DO NOT add any CDNs to Global Warn or Approve unless you want it to apply to JS as well.
     */


    table.rows.push( createAboutRow('2022-05-12',"1.0.0.05","#33, #20 - special chars in headings, #28 - scroll", showRepoLinks === true ? repoLinks : null ) );
    
    table.rows.push( createAboutRow('2022-05-09',"1.0.0.04","#21, #22, #23", showRepoLinks === true ? repoLinks : null ) );
    table.rows.push( createAboutRow('2022-05-09',"1.0.0.04","#7, #9, #10, #11, #13, #16, #17, #18, #19", showRepoLinks === true ? repoLinks : null ) );
    table.rows.push( createAboutRow('2022-05-05',"1.0.0.03","#5 - Add FPS Banner ", showRepoLinks === true ? repoLinks : null ) );
    table.rows.push( createAboutRow('2022-05-03',"1.0.0.02","#2, #3, initial test release", showRepoLinks === true ? repoLinks : null ) );
    
    return { table: table };

}

export function createAboutRow( date: string, version: string, focus: any, repoLinks: IRepoLinks | null ) {

    let fullFocus = convertIssuesMarkdownStringToSpan( focus, repoLinks );

    let tds = [<span style={{whiteSpace: 'nowrap'}} >{ date }</span>, 
        <span style={{whiteSpace: 'nowrap'}} >{ version }</span>, 
        <span>{ fullFocus }</span>,] ;

    return tds;
}