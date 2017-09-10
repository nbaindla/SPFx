declare interface IReactCustomPropertyPaneStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel:string;
  DescriptionFieldLabel: string;
  SelectList: string;
  SelectView : string;
  SelectLibrary: string;
  GroupList: string;
  GroupLibraries: string;
  PageDataSource: string;
  PageDesign: string;
  ddlListView : string;
}

declare module 'reactCustomPropertyPaneStrings' {
  const strings: IReactCustomPropertyPaneStrings;
  export = strings;
}
