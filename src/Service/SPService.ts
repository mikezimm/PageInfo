import { INavLink } from 'office-ui-fabric-react/lib/Nav';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import { HTMLRegEx, IHTMLRegExKeys, IRegExTag } from './htmlTags';
//This should go into npmFunctions v1.0.231 ish
// export const RegexHeading14StartG = HTMLRegEx.h14.openG;
// export const RegexHeading14StartN = /<h[1-4](.*?)>/;
// export const RegexHeading14EndG = /<\/h[1-4]>/g;

// export const RegexHeading13StartG = /<h[1-3](.*?)>/g;
// export const RegexHeading13StartN = /<h[1-3](.*?)>/;
// export const RegexHeading13EndG = /<\/h[1-3]>/g;

// export const RegexHeading12StartG = /<h[1-2](.*?)>/g;
// export const RegexHeading12StartN = /<h[1-2](.*?)>/;
// export const RegexHeading12EndG = /<\/h[1-2]>/g;



export class SPService {
  /* Array to store all unique anchor URLs */
  private static allUrls: string[] = [];

  /**
   * Returns the unique Anchor URL for a heading
   * @param headingValue The text value of the heading
   * @returns anchorUrl
   */
  private static GetAnchorUrl(headingValue: string): string {
    let urlExists = true;

    // Great catch @mikezimm, I included all the extra chars you suggested.
    // Is it a prior function or regex are replacing the & with empty string before that point so your line 21 is not finding it?
    // So, for this part. The problem with line 21 not finding the & char is due to line 20 removing the char. With adding the & char to line 20 we exclude it from our regex expression. .replace(/[^a-zA-Z0-9.,()\-& ]/g, "")

    // .replace(/[^a-zA-Z0-9.,()\-& ]/g, "") replaces chars except a - z, 0 - 9 , & ( ) and a . with ""
    // .replace(/'|?|\|/| |&/g, "-") replaces any blanks and special characters (list is for sure not complete) with "-"
    // .replace(/--+/g, "-") replaces any additional - with only one -; e.g. --- get replaced with -, -- get replaced with - etc.
    let anchorUrl = `#${headingValue
      .replace(/[^a-zA-Z0-9.,()\-& ]/g, "") //https://github.com/mikezimm/PageInfo/issues/20
      .replace(/\'|\?|\\|\/| |\&/g, "-")
      .replace(/--+/g, "-")}`.toLowerCase();
    let urlSuffix = 1;
    while (urlExists === true) {
      urlExists = (this.allUrls.indexOf(anchorUrl) === -1) ? false : true;
      if (urlExists) {
        anchorUrl = anchorUrl + `-${urlSuffix}`;
        urlSuffix++;
      }
    }
    return anchorUrl;
  }
  
  /**
   * Returns the decoded html string
   * @param input the html string
   * @returns decoded string
   */
  private static htmlDecode(input: string) {
    var doc = new DOMParser().parseFromString(input, "text/html");
    return doc.documentElement.textContent;
  }


  


  /**
   * Returns the Anchor Links for Nav element
   * @param context Web part context
   * @returns anchorLinks
   */

  public static async GetAnchorLinks(context: WebPartContext, anchors: IHTMLRegExKeys = 'h14' ) {

    //This gets all the required regex expressions for finding the requested anchors
    const regObj :IRegExTag = HTMLRegEx[ anchors ];

    const anchorLinks: INavLink[] = [];

    try {
      /* Page ID on which the web part is added */
      const pageId = context.pageContext.listItem.id;

      /* Get the canvasContent1 data for the page which consists of all the HTML */
      const data = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${pageId})`, SPHttpClient.configurations.v1);
      const jsonData = await data.json();
      const canvasContent1 = jsonData.CanvasContent1;
      const canvasContent1JSON: any[] = JSON.parse(canvasContent1);

      /* Initialize variables to be used for sorting and adding the Navigation links */
      let headingIndex = 0;
      let subHeadingIndex = -1;
      let headingOrder = 0;
      let prevHeadingOrder = 0;

      /* Traverse through all the Text web parts in the page */
      canvasContent1JSON.map((webPart) => {
        if (webPart.innerHTML) {
          let HTMLString: string = webPart.innerHTML;

          while (HTMLString.search(regObj.openG) !== -1) {
            const lengthFirstOccurence = HTMLString.match(regObj.openG)[0].length;
            /* The Header Text value */
            const headingValue = this.htmlDecode(HTMLString.substring(HTMLString.search(regObj.openG) + lengthFirstOccurence, HTMLString.search(regObj.closeG)));

            headingOrder = parseInt(HTMLString.charAt(HTMLString.search(regObj.openG) + 2));

            const anchorUrl = this.GetAnchorUrl(headingValue);
            this.allUrls.push(anchorUrl);

            /* Add links to Nav element */
            if (anchorLinks.length === 0) {
              anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
            } else {
              if (headingOrder <= prevHeadingOrder) {
                /* Adding or Promoting links */
                switch (headingOrder) {
                  case 2:
                    anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                    headingIndex++;
                    subHeadingIndex = -1;
                    break;
                  case 4:
                    if (subHeadingIndex > -1) {
                      anchorLinks[headingIndex].links[subHeadingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                    } else {
                      anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                    }
                    break;
                  default:
                    anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                    subHeadingIndex = anchorLinks[headingIndex].links.length - 1;
                    break;
                }
              } else {
                /* Making sub links */
                if (headingOrder === 3) {
                  anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                  subHeadingIndex = anchorLinks[headingIndex].links.length - 1;
                } else {
                  if (subHeadingIndex > -1) {
                    anchorLinks[headingIndex].links[subHeadingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                  } else {
                    anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                  }
                }
              }
            }
            prevHeadingOrder = headingOrder;

            /* Replace the added header links from the string so they don't get processed again */
            HTMLString = HTMLString.replace(regObj.open, '').replace(`</h${headingOrder}>`, '');
          }
        }
      });
    } catch (error) {
      console.log(error);
    }

    console.log('FPS Page Info AnchorLinks', anchorLinks);
    return anchorLinks;
  }
}