import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';

export interface IAdvancedPagePropertiesProps {
  showSomeProps: boolean;
  showOOTBProps: boolean;
  context: WebPartContext;
  title: string;
  defaultExpanded: boolean;
  selectedProperties: string[];
  themeVariant: IReadonlyTheme | undefined;

  propsStyle: React.CSSProperties;

}
