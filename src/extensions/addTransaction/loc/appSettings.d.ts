declare interface IAppSettings {
    ProjectTaskListName:"Project Task";
    TaskTransactionListName:"Project Transaction"
  }
  
  declare module 'appSettings' {
    const appSettings: IAppSettings;
    export = appSettings;
  }