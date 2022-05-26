# Fetch ADO (Azure DevOps) Data

This script copies the contents of Azure DevOps Boards to another project.

It uses Azure Functions and Azure Tables.

## Create destination project

Create the destination project.

At this time, be sure to align the **Process (Agile, Scrum, etc.) with the source project**.

Create a team named `Copy` in the destination project.  
(There should be a checkbox asking if you want to create an Area at the same time, but leave it ON.)

## Usage

The code is written in TypeScript.

First, you need a Node.js environment. The author's environment is `v16.14.2`.

Install the necessary global packages.

```shell-session
npm i -g yarn typescript ts-node typesync
```

Install libraries as project root.

```shell-session
yarn
```

Install the VSCode extensions.

- Azure Tools
- Azurite

If you want to edit the code, you may also want to include the following extensions

- ESLint
- Prettier

## local configuration

On the VSCode, you can use [./CopyProject/index.ts](./CopyProject/index.ts) and press `F5`.

The debugger will start, but it will mock up an error such as no environment variable is entered.

If it doesn't even start, I have a theory that you need to install `azure-functions-core-tools` (I have a feeling it will work without it).

As soon as you mock up, `local.settings.json` should be generated in the project root.

Fill in the environment variable information. `{Source, Destination}Password` is your Personal Access Token.

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "AzureStorageConnectionString": "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;TableEndpoint=http://127.0.0.1:10002/devstoreaccount1",
    "TableName": "",
    "SourceUsername": "",
    "SourcePassword": "",
    "SourceOrganization": "",
    "SourceProject": "",
    "SourceTeam": "",
    "DestinationUsername": "",
    "DestinationPassword": "",
    "DestinationOrganization": "",
    "DestinationProject": "",
    "DestinationTeam": ""
  }
}
```

## Mock Azure Storage locally

Mock Azure Storage as it is required to run.

Press `Ctrl`+`Shift`+`P`, type `Azurite: Start Table Service` and run.

## Run locally

Press `F5` on `index.ts` again.

Go to the `Azure Tools` extension screen from the Azure icon on the left side of the screen.

Find `CopyProject` in `Local Project` in `WORKSPACE` and right click on it.

Click `Execute Functions Now`. The request body is not loaded, so it can be anything.

It should work.

## If it works

Prepare Azure Functions and Azure Tables for production deployment.
