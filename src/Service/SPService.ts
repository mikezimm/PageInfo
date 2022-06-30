import { INavLink } from 'office-ui-fabric-react/lib/Nav';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { HTMLRegEx, IHTMLRegExKeys, IRegExTag } from './htmlTags';

export class SPService {
  /* Array to store all unique anchor URLs */
  private static allUrls: string[] = [];

  /**
   * Returns the unique Anchor URL for a heading
   * @param headingValue The text value of the heading
   * @returns anchorUrl
   */
   private static GetAnchorUrl(headingValue: string = 'Empty Heading'): string {
    let anchorUrl = `#${headingValue
      .toLowerCase()
      .replace(/[{}|\[\]\<\>#@"'^%`?;:\/=~\\\s\s+]/g, " ")
      .replace(/^(-|\s)*|(-|\s)*$/g, "")
      .replace(/\'|\?|\\|\/| |\&/g, "-")
      .replace(/-+/g, "-")
      .substring(0, 128)
    }`;

    let counter = 1;
    this.allUrls.forEach(url => {
      if (url === anchorUrl) {
        if (counter != 1) {
          anchorUrl = anchorUrl.slice(0, -((counter - 1).toString().length + 1)) + '-' + counter;

        } else {
          anchorUrl += '-1';
        }

        counter++;
      }
    });

    return anchorUrl;
  }

  /**
   * Returns the Anchor Links for Nav element
   * @param context Web part context
   * @param anchors //2022-06-28:  MZ Added this adjustment to accomodate for more flexible use
   * @returns anchorLinks
   */
  public static async GetAnchorLinks(context: WebPartContext, anchors: IHTMLRegExKeys = 'h14' ) {
    const anchorLinks: INavLink[] = [];

    //2022-06-28:  MZ Added this adjustment to accomodate for more flexible use
    let querySelectList = 'h1, h2, h3, h4';
    switch ( anchors ) {

      case 'h1': querySelectList = 'h1'; break;
      case 'h2': querySelectList = 'h2'; break;
      case 'h3': querySelectList = 'h3'; break;
      case 'h4': querySelectList = 'h4'; break;

      case 'h12': querySelectList = 'h1, h2'; break;
      case 'h13': querySelectList = 'h1, h2, h3'; break;
      case 'h14': querySelectList = 'h1, h2, h3, h4'; break;

      default: querySelectList = 'h1, h2, h3, h4';

    }

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
      let prevHeadingOrder = 0;

      this.allUrls = [];

      /* Traverse through all the Text web parts in the page */
      canvasContent1JSON.map((webPart) => {

        // V V V V Added per sample update https://github.com/pnp/sp-dev-fx-webparts/pull/2804/commits/0e09e2a0b879ecf8d889540e1eba4cf23363ae72
        if (webPart.zoneGroupMetadata) {
          const headingValue = webPart.zoneGroupMetadata.displayName;
          const anchorUrl = this.GetAnchorUrl(headingValue);
          this.allUrls.push(anchorUrl);

          /* Add link to Nav element */
          anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: webPart.zoneGroupMetadata.isExpanded });
        }
        // ^ ^ ^ ^ Added per sample update https://github.com/pnp/sp-dev-fx-webparts/pull/2804/commits/0e09e2a0b879ecf8d889540e1eba4cf23363ae72

        if (webPart.innerHTML) {
          const HTMLString: string = webPart.innerHTML;

          const htmlObject = document.createElement('div');
          htmlObject.innerHTML = HTMLString;

          const headers = htmlObject.querySelectorAll( querySelectList );

          headers.forEach(header => {
            const headingValue = header.textContent;

            // V V V V Added per sample update https://github.com/pnp/sp-dev-fx-webparts/pull/2804/commits/0e09e2a0b879ecf8d889540e1eba4cf23363ae72
            let headingOrder = parseInt(header.tagName.substring(1));

            if (webPart.zoneGroupMetadata) {
              headingOrder++;
            }
            // ^ ^ ^ ^ Added per sample update https://github.com/pnp/sp-dev-fx-webparts/pull/2804/commits/0e09e2a0b879ecf8d889540e1eba4cf23363ae72

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
          });
        }
      });
    } catch (error) {
      console.log(error);
    }

    console.log('FPS Page Info AnchorLinks', anchorLinks);

    return anchorLinks;
  }
}