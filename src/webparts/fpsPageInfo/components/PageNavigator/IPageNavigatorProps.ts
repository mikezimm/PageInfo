import { INavLink } from 'office-ui-fabric-react/lib/Nav';

import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';

export type IMinHeading = 'h3' | 'h2' | 'h1' ;

export interface IPageNavigatorProps {
  description: string;
  showTOC: boolean;
  tocExpanded: boolean;
  minHeadingToShow: IMinHeading;
  anchorLinks: INavLink[];
  themeVariant: IReadonlyTheme | undefined;

  tocStyle: React.CSSProperties;
}
