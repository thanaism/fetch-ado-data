import axios from 'axios';
import { destinationProfile, ProjectProfile, sourceProfile } from '../environment/profile';
import { User } from '../users/user';

export class DestinationComment {
  private static profile = destinationProfile;

  public static async get(workItemId: number): Promise<Comment[]> {
    const response = await CommentAPI.get(this.profile, workItemId);
    return response.data.comments;
  }

  public static async add(workItemId: number, commentText: string) {
    const response = await CommentAPI.add(this.profile, workItemId, commentText);
    return response.data;
  }

  public static update(workItemId: number, commentId: number, commentText: string) {
    return CommentAPI.update(this.profile, workItemId, commentId, commentText);
  }
}

export class SourceComment {
  private static profile = sourceProfile;

  public static async get(workItemId: number): Promise<Comment[]> {
    const response = await CommentAPI.get(this.profile, workItemId);
    return response.data.comments;
  }
}

class CommentAPI {
  public static add(profile: ProjectProfile, workItemId: number, commentText: string) {
    const { organization, project, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/` +
      `_apis/wit/workItems/${workItemId}/comments`;
    return axios.post<Comment>(
      url,
      { text: commentText },
      {
        auth: { username: username, password: password },
        params: { 'api-version': '7.1-preview.3' },
      },
    );
  }

  public static get(profile: ProjectProfile, workItemId: number) {
    const { organization, project, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/` +
      `_apis/wit/workItems/${workItemId}/comments`;
    return axios.get<CommentGetAPIResponse>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '7.1-preview.3' },
    });
  }

  public static update(
    profile: ProjectProfile,
    workItemId: number,
    commentId: number,
    commentText: string,
  ) {
    const { organization, project, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/` +
      `_apis/wit/workItems/${workItemId}/comments/${commentId}`;
    return axios.patch<Comment>(
      url,
      { text: commentText },
      {
        auth: { username: username, password: password },
        params: { 'api-version': '6.0-preview.3' },
      },
    );
  }
}

type CommentGetAPIResponse = {
  totalCount: number;
  count: number;
  comments: Comment[];
};

type Comment = {
  workItemId: number;
  id: number;
  version: number;
  text: string;
  createdBy: User;
  createdDate: string;
  modifiedBy: User;
  modifiedDate: string;
  url: string;
};
