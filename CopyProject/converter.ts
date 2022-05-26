import { attachmentDownloadAndUpload } from './attachments/attachmentDuplicator';
import { Memo } from './keyValueStore/memo';
import { destinationProfile } from './environment/profile';
import { DestinationWorkItem } from './workItems/workItemAPI';
import { getLogger } from 'log4js';

const logger = getLogger();

export const addMetadata = (content: string, author: string, createdDate: string): string => {
  return (
    `<table style="border: 1px solid #333;">` +
    `<thead style=" background-color: #333;color: #fff;">` +
    `<tr><th colspan="2">Metadata</th></tr></thead>` +
    `<tbody><tr><td style="border: 1px solid #333;">Author: ${author}</td>` +
    `<td style="border: 1px solid #333;">CreatedDate: ${createdDate}</td>` +
    `</tr></tbody></table>${content}`
  );
};

export const convertLinks = async (content: string): Promise<string> => {
  logger.debug(`convertLinks ---`);
  const regExpWorkItem = RegExp(
    `(<a[^>]*href=")` + // 0
      `(https://dev\.azure\.com/[^/]+/[^/]+/_workitems/edit/[0-9]+)` + // 1
      `("[^>]* data-vss-mention="version:1\.0"[^>]*>#)` + // 2
      `([1-9][0-9]*)(</a>)`, // 3
    'g',
  );
  const memo = await Memo.build();
  const contentWorkItemReplaced = content.replace(regExpWorkItem, (match, ...p) => {
    const capturedId = p[3] as string;
    const counterpartId = memo.has('counterpartId', capturedId)
      ? memo.get('counterpartId', capturedId)
      : 0;
    const url = toUrl(counterpartId as number);
    return [p[0], url, p[2], counterpartId, p[4]].join('');
  });

  logger.trace(`contentWorkItemReplaced:\n${contentWorkItemReplaced}`);

  const regExpAttachment = RegExp(
    `(<img[^>]*src=")` + // 0
      `(https://dev\.azure\.com/[^/]+/[^/]+/_apis/wit/attachments/)` + // 1
      `([^?]+)` + // 2
      `(\\?fileName=[^"]*)` + // 3
      `("[^>]*>)`, // 4
    'g',
  );

  const matches = [...contentWorkItemReplaced.matchAll(regExpAttachment)];

  if (matches.length === 0) {
    logger.debug(`No attachments included.`);
    return contentWorkItemReplaced;
  }

  logger.trace(`matches:\n\t${JSON.stringify(matches)}`);

  for (const match of matches) {
    const sourceAttachmentId = match[3];
    logger.trace(`sourceAttachmentId: ${sourceAttachmentId}`);

    const alreadyDupulicated = memo.has('counterpartUrl', sourceAttachmentId);

    if (!alreadyDupulicated) {
      const response = await attachmentDownloadAndUpload(
        sourceAttachmentId,
        encodeURIComponent('image.png'),
      );
      memo.upsert('counterpartUrl', sourceAttachmentId, response.url);
      logger.debug(`Image Uploaded: ${response.url}`);
    } else {
      logger.debug(`Attachment already dupulicated: ${sourceAttachmentId}`);
    }
  }

  const contentFullyReplaced: string = contentWorkItemReplaced.replace(
    regExpAttachment,
    (_, ...p) => {
      logger.trace(`p[2]: ${String(p[2])}`);
      logger.trace(`p[3]: ${String(p[3])}`);
      const sourceAttachmentId = p[2] as string;
      const duplicatedImageUrl = memo.get('counterpartUrl', sourceAttachmentId);
      return [p[0], duplicatedImageUrl, p[4]].join('');
    },
  );

  logger.trace(`contentFullyReplaced:\n${contentFullyReplaced}`);

  return contentFullyReplaced;
};

const toUrl = async (workItemId: number) => {
  try {
    const response = await DestinationWorkItem.get(String(workItemId));
    return response._links.html.href;
  } catch {
    const { organization, project } = destinationProfile;
    return `https://dev.azure.com/${organization}/${project}/_workitems/recentlyupdated/`;
  }
};
