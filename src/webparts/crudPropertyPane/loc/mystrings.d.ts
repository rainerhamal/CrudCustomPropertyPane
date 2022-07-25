declare interface ICrudPropertyPaneWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  WorkItemFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'CrudPropertyPaneWebPartStrings' {
  const strings: ICrudPropertyPaneWebPartStrings;
  export = strings;
}
