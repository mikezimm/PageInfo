import { IAdvancedPagePropertiesProps } from "./AdvPageProps/components/IAdvancedPagePropertiesProps";
import { IPageNavigatorProps } from "./PageNavigator/IPageNavigatorProps";

export interface IFpsPageInfoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  pageNavigator: IPageNavigatorProps;

  advPageProps: IAdvancedPagePropertiesProps;

}
