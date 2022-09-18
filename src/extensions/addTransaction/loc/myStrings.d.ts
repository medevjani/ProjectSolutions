declare interface IAddTransactionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'AddTransactionCommandSetStrings' {
  const strings: IAddTransactionCommandSetStrings;
  export = strings;
}
