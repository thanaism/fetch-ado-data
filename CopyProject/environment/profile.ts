import { EnvironmentVariables } from './env';

export type ProjectProfile = {
  organization: string;
  project: string;
  team: string;
  username: string;
  password: string;
};

const environmentVariables = EnvironmentVariables.instance;

export const sourceProfile: ProjectProfile = {
  organization: encodeURIComponent(environmentVariables.SourceOrganization),
  project: encodeURIComponent(environmentVariables.SourceProject),
  team: encodeURIComponent(environmentVariables.SourceTeam),
  username: environmentVariables.SourceUsername,
  password: environmentVariables.SourcePassword,
};

export const destinationProfile: ProjectProfile = {
  organization: encodeURIComponent(environmentVariables.DestinationOrganization),
  project: encodeURIComponent(environmentVariables.DestinationProject),
  team: encodeURIComponent(environmentVariables.DestinationTeam),
  username: environmentVariables.DestinationUsername,
  password: environmentVariables.DestinationPassword,
};
