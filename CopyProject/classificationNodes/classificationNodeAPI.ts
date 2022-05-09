import axios from 'axios';
import { ProjectProfile, sourceProfile, destinationProfile } from '../environment/profile';

export class SourceClassificationNode {
  private static profile = sourceProfile;

  public static async getRootNodes(): Promise<ClassificationNode[]> {
    const response = await ClassificationNodeAPI.getRootNodes(this.profile);
    return response.data.value;
  }
}

export class DestinationClassificationNode {
  private static profile = destinationProfile;

  public static upsert(structureGroup: StructureGroup, path: string, name: string) {
    return ClassificationNodeAPI.upsert(this.profile, structureGroup, path, name);
  }

  public static async getRootNodes(): Promise<
    [area: ClassificationNode, iteration: ClassificationNode]
  > {
    const rootNodes = (await ClassificationNodeAPI.getRootNodes(this.profile)).data.value;
    const pseudoRootName = destinationProfile.team;

    const _areaRoot: ClassificationNode | undefined = rootNodes
      .find((node) => node.structureType === 'area')
      .children?.find((node) => node?.name === pseudoRootName);

    const areaRoot: ClassificationNode =
      _areaRoot ?? (await this.upsert('Areas', '', pseudoRootName)).data;

    const _iterationRoot: ClassificationNode | undefined = rootNodes
      .find((node) => node.structureType === 'iteration')
      .children?.find((node) => node?.name === pseudoRootName);

    const iterationRoot: ClassificationNode =
      _iterationRoot ?? (await this.upsert('Iterations', '', pseudoRootName)).data;

    return [areaRoot, iterationRoot];
  }
}

class ClassificationNodeAPI {
  public static getRootNodes(profile: ProjectProfile) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/classificationnodes`;
    return axios.get<ClassificationNodeGetResponse>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '5.0', $depth: '100' },
    });
  }

  public static upsert(
    profile: ProjectProfile,
    structureGroup: StructureGroup,
    path: string,
    name: string,
  ) {
    const { organization, project, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/` +
      `_apis/wit/classificationnodes/${structureGroup}/${path}`;
    return axios.post<ClassificationNode>(
      url,
      { name: name },
      {
        auth: { username: username, password: password },
        params: { 'api-version': '6.0' },
      },
    );
  }
}

export type StructureGroup = 'Areas' | 'Iterations';
export type StructureType = 'area' | 'iteration';
export type ClassificationNode = {
  id: number;
  identifier: string;
  name: string;
  structureType: StructureType;
  hasChildren: boolean;
  children: ClassificationNode[];
  path: string;
  _links: {
    self: {
      href: string;
    };
    parent: {
      href: string;
    };
  };
  url: string;
};

type ClassificationNodeGetResponse = {
  count: number;
  value: ClassificationNode[];
};
