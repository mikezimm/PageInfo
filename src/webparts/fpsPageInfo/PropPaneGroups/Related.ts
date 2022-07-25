import {
  //  IPropertyPanePage,
   IPropertyPaneGroup,
  //  PropertyPaneLabel,
  //  IPropertyPaneLabelProps,
  //  PropertyPaneHorizontalRule,
    PropertyPaneTextField, IPropertyPaneTextFieldProps,
  //   PropertyPaneLink, IPropertyPaneLinkProps,
  // PropertyPaneDropdown, IPropertyPaneDropdownProps,
  IPropertyPaneDropdownOption,PropertyPaneToggle,
  IPropertyPaneField,
  //  IPropertyPaneConfiguration,
  //  PropertyPaneButton,
  //  PropertyPaneButtonType,
  //   PropertyPaneSlider, IPropertyPaneSliderProps,
  // PropertyPaneHorizontalRule,
  // PropertyPaneSlider
} from '@microsoft/sp-property-pane';

import { IRelatedItemsProps, IRelatedKey } from '@mikezimm/npmfunctions/dist/RelatedItems/IRelatedItemsProps';

// import { getHelpfullErrorV2 } from '../Logging/ErrorHandler';
// import { JSON_Edit_Link } from './zReusablePropPane';

// import * as strings from 'FpsPageInfoWebPartStrings';
import { IFpsPageInfoWebPartProps } from '../IFpsPageInfoWebPartProps';

export function buildRelatedProps( wpProps: IFpsPageInfoWebPartProps, name: IRelatedKey ) {

  var groupFields: IPropertyPaneField<any>[] = [];

  groupFields.push(PropertyPaneToggle(`${name}showItems`, {
    label: "Add Related Items",
    onText: "On",
    offText: "Off",
    // disabled: true,
  }));

  groupFields.push(PropertyPaneTextField(`${name}heading`, {
    label: 'Heading - accordion',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`${name}web`, {
    label: 'Url to site - starting with /sites/...',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`${name}listTitle`, {
    label: 'Title of the list or library which has related items',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`${name}restFilter`, {
    label: 'Rest filter - click bright yellow icon for examples',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`${name}displayProp`, {
    label: 'Static field name of Related item Label',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`${name}linkProp`, {
    label: 'Static field name of Related item Link',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneToggle(`${name}isExpanded`, {
    label: "Expand by default",
    onText: "On",
    offText: "Off",
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  groupFields.push(PropertyPaneTextField(`relatedStyle`, {
    label: 'React.CSS Item Styles',
    disabled: wpProps[`${name}showItems`] === false ? true : false,
  }));

  const RelatedGroup: IPropertyPaneGroup = {
    groupName: `Related Props ${name.replace(/\D/g, '')}`,
    isCollapsed: true,
    groupFields: groupFields
  };

  return RelatedGroup;

}
