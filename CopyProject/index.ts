// eslint-disable-next-line import/no-unresolved
import { AzureFunction, Context, HttpRequest } from '@azure/functions';
import { EnvironmentVariables } from './environment/env';
import { duplicateClassificationNodes } from './classificationNodes/classificationNodeDuplicator';
import { duplicateWorkItems } from './workItems/workItemDuplicator';
import { Memo } from './keyValueStore/memo';
import HTTP_STATUS_CODES from 'http-status-enum';
import { getLogger } from 'log4js';
import { AxiosError } from 'axios';

const logger = getLogger();
logger.level = process.env.NODE_ENV !== 'production' ? 'debug' : 'error';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const httpTrigger: AzureFunction = async (context: Context, req: HttpRequest): Promise<void> => {
  logger.debug('HTTP trigger function processed a request.');

  try {
    EnvironmentVariables.instance;
    const memo = await Memo.build();
    await duplicateClassificationNodes();
    await duplicateWorkItems();
    logger.warn(`Duplication succeeded. Sync HashMap to AzureTables...`);
    await memo.sync();
    const responseMessage = 'yes';
    context.res = { status: 200, body: responseMessage };
  } catch (e: unknown) {
    if (e instanceof Error) {
      logger.error(e.message);
      context.res = { status: HTTP_STATUS_CODES.INTERNAL_SERVER_ERROR, body: e.message };
      context.done();
    }
    if (e instanceof AxiosError) {
      logger.error(e.response);
    }
  }
};

export default httpTrigger;
