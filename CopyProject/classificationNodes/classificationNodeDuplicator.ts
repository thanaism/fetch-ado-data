import {
  ClassificationNode,
  DestinationClassificationNode,
  SourceClassificationNode,
  StructureGroup,
  StructureType,
} from './classificationNodeAPI';
import { EnvironmentVariables } from '../environment/env';
import { DestinationIteration } from '../iterations/iterationAPI';

export const duplicateClassificationNodes = async (): Promise<void> => {
  const sourceRootNodes = await SourceClassificationNode.getRootNodes();
  const destinationRootNodes = await DestinationClassificationNode.getRootNodes();

  const identifiers: string[] = destinationRootNodes.flatMap((node) =>
    node.structureType === 'iteration' ? [node.identifier] : [],
  );

  identifiers.push(
    ...(await duplicateChildrenIfNotExists(
      toChildren(sourceRootNodes),
      toChildren(destinationRootNodes),
    )),
  );

  await addIterationsToDestinationTeamIfNotExists(identifiers);
};

const addIterationsToDestinationTeamIfNotExists = async (identifiers: string[]) => {
  // In the Azure DevOps Rest API,
  // the contents in the "id" and "identifier" are different depending on the API,
  // even for iteration objects that are supposed to be of the same type.
  const existingIterationIdentifiers: string[] = (await DestinationIteration.list()).data.value.map(
    (iteration) => iteration.id,
  );

  const iterationIdentifiersToAdd: string[] = identifiers.filter(
    (identifier) => !existingIterationIdentifiers.includes(identifier),
  );

  for (const identifier of iterationIdentifiersToAdd)
    await DestinationIteration.addToTeam(identifier);
};

const toChildren = (nodes: ClassificationNode[]): ClassificationNode[] =>
  nodes.flatMap((node) => (node?.hasChildren ? node.children : []));

const duplicateChildrenIfNotExists = async (
  sourceChildren: ClassificationNode[],
  copyChildren: ClassificationNode[],
): Promise<string[]> => {
  const identifiers: string[] = [];
  for (const sourceChild of sourceChildren) {
    const counterpart = copyChildren.find(
      (copyChild) =>
        copyChild?.name === sourceChild.name &&
        copyChild?.structureType === sourceChild.structureType,
    );

    const identifier =
      counterpart != null
        ? counterpart.identifier
        : (
            await DestinationClassificationNode.upsert(
              toStructureGroup(sourceChild.structureType),
              toPathRoot(sourceChild.path),
              sourceChild.name,
            )
          ).data.identifier;

    if (sourceChild.structureType === 'iteration') identifiers.push(identifier);

    identifiers.push(
      ...(await duplicateChildrenIfNotExists(
        sourceChild.hasChildren ? sourceChild.children : [],
        counterpart?.hasChildren ? counterpart.children : [],
      )),
    );
  }
  return identifiers;
};

const toStructureGroup = (type: StructureType): StructureGroup =>
  type === 'area' ? 'Areas' : 'Iterations';

const toPathRoot = (path: string) =>
  [EnvironmentVariables.instance.DestinationTeam].concat(path.split('\\').slice(3, -1)).join('\\');
