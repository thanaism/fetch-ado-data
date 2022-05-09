// eslint-disable-next-line import/no-unresolved
export class EnvironmentVariables {
  AzureStorageConnectionString: string;
  TableName: string;
  SourceUsername: string;
  SourcePassword: string;
  SourceOrganization: string;
  SourceProject: string;
  SourceTeam: string;
  DestinationUsername: string;
  DestinationPassword: string;
  DestinationOrganization: string;
  DestinationProject: string;
  DestinationTeam: string;

  private static _instance: EnvironmentVariables;

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  private constructor() {}

  public static get instance(): EnvironmentVariables {
    if (this._instance == null) {
      this._instance = new EnvironmentVariables();
      /* eslint-disable @typescript-eslint/no-non-null-assertion */
      this._instance.AzureStorageConnectionString = process.env.AzureStorageConnectionString!;
      this._instance.TableName = process.env.TableName!;
      this._instance.SourceUsername = process.env.SourceUsername!;
      this._instance.SourcePassword = process.env.SourcePassword!;
      this._instance.SourceOrganization = process.env.SourceOrganization!;
      this._instance.SourceProject = process.env.SourceProject!;
      this._instance.SourceTeam = process.env.SourceTeam!;
      this._instance.DestinationUsername = process.env.DestinationUsername!;
      this._instance.DestinationPassword = process.env.DestinationPassword!;
      this._instance.DestinationOrganization = process.env.DestinationOrganization!;
      this._instance.DestinationProject = process.env.DestinationProject!;
      this._instance.DestinationTeam = process.env.DestinationTeam!;
      /* eslint-enable @typescript-eslint/no-non-null-assertion */
      this._instance.validate();
    }
    return this._instance;
  }

  private validate(): void {
    const emptyKey = Object.getOwnPropertyNames(this).find((key) => this[key] == null);
    if (emptyKey == null) return;
    const errorMessage = `Environment variable "${emptyKey}" is undefined.`;
    throw Error(errorMessage);
  }
}
