/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { TableClient, TableServiceClient } from '@azure/data-tables';

export class AzureTables {
  private static azureTables: AzureTables;
  private tableClient: TableClient;

  private constructor() {
    this.tableClient = TableClient.fromConnectionString(
      process.env.AzureStorageConnectionString!,
      process.env.TableName!,
      { allowInsecureConnection: process.env.NODE_ENV !== 'production' },
    );
  }

  public static async build() {
    if (this.azureTables == null) {
      const tableServiceClient = TableServiceClient.fromConnectionString(
        process.env.AzureStorageConnectionString!,
        { allowInsecureConnection: process.env.NODE_ENV !== 'production' },
      );
      await tableServiceClient.createTable(process.env.TableName!);
      this.azureTables = new AzureTables();
    }
    return this.azureTables;
  }

  public listEntities() {
    return this.tableClient.listEntities();
  }

  public async getEntity(partitionKey: string, rowKey: string) {
    return this.tableClient.getEntity(partitionKey, rowKey);
  }

  public async upsertEntity(
    partitionKey: string,
    rowKey: string,
    value: { [key: string]: unknown },
  ) {
    const entity = { partitionKey: partitionKey, rowKey: rowKey, ...value };
    await this.tableClient.upsertEntity(entity);
  }
}
