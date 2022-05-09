import axios from 'axios';
import { destinationProfile, ProjectProfile, sourceProfile } from '../environment/profile';

export class SourceAttachment {
  private static profile = sourceProfile;

  public static async get(id: string) {
    const response = await AttachmentAPI.get(this.profile, id);
    return response.data;
  }
}

export class DestinationAttachment {
  private static profile = destinationProfile;

  public static async create(filename: string, blob: string) {
    const response = await AttachmentAPI.create(this.profile, filename, blob);
    return response.data;
  }
}

class AttachmentAPI {
  public static get(profile: ProjectProfile, id: string) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/attachments/${id}`;
    return axios.get<string>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0' },
      responseType: 'arraybuffer',
    });
  }

  public static create(profile: ProjectProfile, filename: string, blob: string) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/attachments/`;
    type attachmentPostResponse = { id: string; url: string };
    return axios.post<attachmentPostResponse>(url, blob, {
      auth: { username: username, password: password },
      params: { fileName: filename, 'api-version': '6.0' },
      headers: { 'Content-Type': 'application/octet-stream' },
    });
  }
}
