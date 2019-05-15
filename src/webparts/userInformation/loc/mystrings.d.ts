declare interface IUserInformationWebPartStrings {
  PropertyPaneDescription: string;
  ShowPhotoTargetProperty: string;
  ShowPhotoLabel: string;  
  ShowPhotoCalloutText:string;
  ShowPhotoOnText:string;
  ShowPhotoOffText:string;
  PersonaSizeTargetProperty:string;
  PersonaSizeLabel:string;
  GraphServiceURI:string;
  GraphServicePhotoURI:string;
}

declare module 'UserInformationWebPartStrings' {
  const strings: IUserInformationWebPartStrings;
  export = strings;
}
