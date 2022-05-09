import { DestinationComment, SourceComment } from './commentAPI';
import { addMetadata, convertLinks } from '../converter';
import { Memo } from '../keyValueStore/memo';
import { WorkItem } from '../workItems/workItemAPI';

export const duplicateComments = async (sourceWorkItem: WorkItem): Promise<void> => {
  const memo = await Memo.build();
  const sourceComments = (await SourceComment.get(sourceWorkItem.id)).reverse();

  const commentsNotYetDuplicated = sourceComments.filter(
    (comment) => !memo.has('counterpartId', comment.id),
  );

  for (const comment of commentsNotYetDuplicated) {
    const response = await DestinationComment.add(
      memo.get('counterpartId', sourceWorkItem.id) as number,
      await convertLinks(
        addMetadata(comment.text, comment.createdBy.displayName, comment.createdDate),
      ),
    );
    memo.upsert('counterpartId', comment.id, response.id);
    memo.upsert('previousChangedDate', comment.id, response.modifiedDate);
  }

  const destinationComments = await DestinationComment.get(
    memo.get('counterpartId', sourceWorkItem.id) as number,
  );

  for (const comment of sourceComments) {
    const counterpartId = memo.get('counterpartId', comment.id);
    const previousChangedDate = memo.get('previousChangedDate', comment.id);

    const counterpart = destinationComments.find(
      (comment) => comment.id === (counterpartId as number),
    );
    if (counterpart == null) throw Error;

    if (Date.parse(previousChangedDate as string) < Date.parse(counterpart.modifiedDate))
      await DestinationComment.update(
        memo.get('counterpartId', sourceWorkItem.id) as number,
        counterpartId as number,
        await convertLinks(
          addMetadata(comment.text, comment.createdBy.displayName, comment.createdDate),
        ),
      );
  }
};
