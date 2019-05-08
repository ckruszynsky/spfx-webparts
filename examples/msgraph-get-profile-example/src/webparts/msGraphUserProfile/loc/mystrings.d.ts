declare interface IMsGraphUserProfileWebPartStrings {
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

declare module 'MsGraphUserProfileWebPartStrings' {
  const strings: IMsGraphUserProfileWebPartStrings;
  export = strings;
}
