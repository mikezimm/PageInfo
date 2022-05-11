
/**
 * These are used to parse heading tags out of html strings
 * FPS Page Info SP Service 
 * NOTE:  h2 in html is really h1 on sharepoint page
 */



//This was based upon examples found in react-page-navigation Service.ts file
// export const RegexHeading14StartG = /<h[1-4](.*?)>/g;
// export const RegexHeading14StartN = /<h[1-4](.*?)>/;
// export const RegexHeading14EndG = /<\/h[1-4]>/g;

// export const RegexHeading13StartG = /<h[1-3](.*?)>/g;
// export const RegexHeading13StartN = /<h[1-3](.*?)>/;
// export const RegexHeading13EndG = /<\/h[1-3]>/g;

// export const RegexHeading12StartG = /<h[1-2](.*?)>/g;
// export const RegexHeading12StartN = /<h[1-2](.*?)>/;
// export const RegexHeading12EndG = /<\/h[1-2]>/g;


 export interface IRegExTag {
    key: string;
    desc: string;
    
    misc?: string;
    tags?: string[];
    open: RegExp;
    close: RegExp;
    openG: RegExp;
    closeG: RegExp;
}



const openTagStr = '<tag(.*?)>';
const closeTagStr = '<\/tag>';

export function createTagRegex ( tag: string, misc: string = '' ) : IRegExTag  {

    let thisReg : IRegExTag= {
        key: tag,
        tags: [ tag ],
        desc: `Finds ${tag} tags in Html`,
        misc: misc,
        open: new RegExp(openTagStr.replace('tag', tag ), ),
        close: new RegExp(closeTagStr.replace('tag', tag ), ),
        openG: new RegExp(openTagStr.replace('tag', tag ), 'g' ),
        closeG: new RegExp(closeTagStr.replace('tag', tag ), 'g' ),
    };

    return thisReg;

}

const h12Tag = 'h[1-2]';
export const h12RegEx :IRegExTag = {
    key: 'h12',
    tags: [ 'h1', 'h2' ],
    desc: 'Finds h1 tags in Html',
    misc: 'Not sure if SharePoint has any h1 on a page.',
    open: new RegExp(openTagStr.replace('tag', h12Tag ), ),
    close: new RegExp(closeTagStr.replace('tag', h12Tag ), ),
    openG: new RegExp(openTagStr.replace('tag', h12Tag ), 'g' ),
    closeG: new RegExp(closeTagStr.replace('tag', h12Tag ), 'g' ),
};

const h13Tag = 'h[1-3]';
export const h13RegEx :IRegExTag = {
    key: 'h13',
    tags: [ 'h1', 'h2', 'h3' ],
    desc: 'Finds h1, h2 and h3 tags in Html',
    misc: 'SharePoint Headings are one index larger... SharePoint Heading1 is really an h2 tag',
    open: new RegExp(openTagStr.replace('tag', h13Tag ), ),
    close: new RegExp(closeTagStr.replace('tag', h13Tag ), ),
    openG: new RegExp(openTagStr.replace('tag', h13Tag ), 'g' ),
    closeG: new RegExp(closeTagStr.replace('tag', h13Tag ), 'g' ),
};

const h14Tag = 'h[1-4]';
export const h14RegEx :IRegExTag = {
    key: 'h14',
    tags: [ 'h1', 'h2', 'h3', 'h4' ],
    desc: 'Finds h1, h2 , h3 and h4 tags in Html',
    misc: 'SharePoint Headings are one index larger... SharePoint Heading1 is really an h2 tag',
    open: new RegExp(openTagStr.replace('tag', h14Tag ), ),
    close: new RegExp(closeTagStr.replace('tag', h14Tag ), ),
    openG: new RegExp(openTagStr.replace('tag', h14Tag ), 'g' ),
    closeG: new RegExp(closeTagStr.replace('tag', h14Tag ), 'g' ),
};

export interface IHTMLRegEx {
    h1?: IRegExTag;
    h2?: IRegExTag;
    h3?: IRegExTag;
    h4?: IRegExTag;
    h12?: IRegExTag;
    h13?: IRegExTag;
    h14?: IRegExTag;
    p?: IRegExTag;
    div?: IRegExTag;
    li?: IRegExTag;
    a?: IRegExTag;
    mark?: IRegExTag;
    bold?: IRegExTag;
    img?: IRegExTag;
    script?: IRegExTag;
    style?: IRegExTag;
}

export type IHTMLRegExKeys = 'h1' | 'h2' | 'h3' | 'h4' | 'h12' | 'h13' | 'h14' | 'p' | 'mark' | 'bold' | 'li' | 'img' ;

export const HTMLRegEx: IHTMLRegEx = {
    h1: createTagRegex( 'h1', 'Finds Not sure in SharePoint Pages'),
    h2: createTagRegex( 'h2', 'Finds Heading1 in SharePoint Pages'),
    h3: createTagRegex( 'h3', 'Finds Heading2 in SharePoint Pages'),
    h4: createTagRegex( 'h4', 'Finds Heading3 in SharePoint Pages'),
    h12: h12RegEx,
    h13: h13RegEx,
    h14: h14RegEx,
    p: createTagRegex( 'p', 'Finds Not sure in SharePoint Pages'),
    div: createTagRegex( 'div', 'Finds Not sure in SharePoint Pages'),
    li: createTagRegex( 'li', 'Finds Not sure in SharePoint Pages'),
    a: createTagRegex( 'a', 'Finds Not sure in SharePoint Pages'),
    mark: createTagRegex( 'mark', 'Finds Not sure in SharePoint Pages'),
    bold: createTagRegex( 'bold', 'Finds Not sure in SharePoint Pages'),
    img: createTagRegex( 'img', 'Finds Not sure in SharePoint Pages'),
    script: createTagRegex( 'script', 'Finds Not sure in SharePoint Pages'),
    style: createTagRegex( 'style', 'Finds Not sure in SharePoint Pages'),


};
