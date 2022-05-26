import axios from 'axios';
import { getLogger } from 'log4js';
import { destinationProfile, ProjectProfile } from '../environment/profile';

const logger = getLogger();

export class DestinationIteration {
  private static profile = destinationProfile;

  public static list() {
    return IterationAPI.list(this.profile);
  }

  public static addToTeam(identifier: string) {
    logger.debug(`Add iteration to destination team: ${identifier}`);
    return IterationAPI.addToTeam(this.profile, identifier);
  }
}

class IterationAPI {
  public static list(profile: ProjectProfile) {
    const { organization, project, team, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/${team}/` +
      `_apis/work/teamsettings/iterations`;
    return axios.get<IterationListAPIResponse>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0' },
    });
  }

  public static addToTeam(profile: ProjectProfile, iterationId: string) {
    const { organization, project, team, username, password } = profile;
    const url =
      `https://dev.azure.com/${organization}/${project}/${team}/` +
      `_apis/work/teamsettings/iterations`;
    return axios.post<Iteration>(
      url,
      { id: iterationId },
      {
        auth: { username: username, password: password },
        params: { 'api-version': '6.0' },
      },
    );
  }
}

type IterationListAPIResponse = {
  count: number;
  value: Iteration[];
};

type Iteration = {
  id: string;
  name: string;
  path: string;
  attributes: IterationAttributes;
  url: string;
};

type IterationAttributes = {
  startDate: string;
  finishDate: string;
  timeFrame: string;
};
