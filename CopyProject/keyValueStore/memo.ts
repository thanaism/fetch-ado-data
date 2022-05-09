import { AzureTables } from './azureTables';

type PartitionKey = 'previousChangedDate' | 'counterpartId' | 'counterpartUrl';

export class Memo {
  private static memo: Memo;
  private azureTables: AzureTables;
  private map: Map<string, Record<string, unknown>>;

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  private constructor() {}

  public static async build() {
    if (this.memo == null) {
      this.memo = new Memo();
      this.memo.map = new Map<string, Record<string, unknown>>();
      this.memo.azureTables = await AzureTables.build();
      const entities = this.memo.azureTables.listEntities();
      for await (const entity of entities) {
        const { partitionKey, rowKey } = entity;
        this.memo.map.set(String([partitionKey, rowKey]), entity);
      }
    }
    return this.memo;
  }

  public has(partitionKey: PartitionKey, rowKey: string | number) {
    return this.map.has(String([partitionKey, rowKey]));
  }
  public get(partitionKey: PartitionKey, rowKey: string | number) {
    return this.map.get(String([partitionKey, rowKey])).value;
  }

  public upsert(partitionKey: PartitionKey, rowKey: string | number, value: unknown) {
    this.map.set(String([partitionKey, rowKey]), { value });
  }

  public async sync() {
    for (const [key, value] of this.map) {
      const [partitionKey, rowKey] = key.split(',');
      await this.azureTables.upsertEntity(partitionKey, rowKey, value);
    }
  }
}
