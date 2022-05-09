import { DestinationAttachment, SourceAttachment } from './attachmentAPI';

export const attachmentDownloadAndUpload = async (
  sourceAttachmentId: string | number,
  encodedFilename: string,
) => {
  const blob = await SourceAttachment.get(String(sourceAttachmentId));
  return DestinationAttachment.create(encodedFilename, blob);
};
