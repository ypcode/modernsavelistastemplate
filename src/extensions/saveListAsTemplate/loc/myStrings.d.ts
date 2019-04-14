declare interface ISaveListAsTemplateCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SaveListAsTemplateCommandSetStrings' {
  const strings: ISaveListAsTemplateCommandSetStrings;
  export = strings;
}
