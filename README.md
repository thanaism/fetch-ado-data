# Fetch ADO (Azure DevOps) Data

This script copies the contents of Azure DevOps Boards to another project.

It uses Azure Functions and Azure Tables.

## Create destination project

Create the destination project.

At this time, be sure to align the **Process (Agile, Scrum, etc.) with the source project**.

Create a team named `Copy` in the destination project.  
(There should be a checkbox asking if you want to create an Area at the same time, but leave it ON.)

## Usage

All code is written in TypeScript and requires a Node.js environment.

The author's environment is `v16.14.2`, but it should work with `v16.0.0` or `v14.17.0` or later.

Install the necessary global packages.

```shell-session
npm i -g yarn typescript ts-node typesync
```

Install libraries as project root.

```shell-session
yarn
```

Install the VSCode extension.

- Azure Tools
- Azurite

If you want to edit the code, you may also want to include the following extensions

- ESLint
- Prettier

## local configuration

For local startup, the Azure Functions Core Tools must be installed.

Install it globally with the following command.

```shell-session
sudo npm i -g azure-functions-core-tools@4 --unsafe-perm true
```

On the VSCode [./CopyProject/index.ts](./CopyProject/index.ts) and press `F5`.

The debugger will start, but it will mock up an error such as no environment variable is entered.

As soon as you mock up, `local.settings.json` should be generated in the project root.

Fill in the environment variable information.

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "AzureStorageConnectionString": "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;TableEndpoint=http://127.0.0.1:10002/devstoreaccount1",
    "TableName": "Azure Tables table name (any value for local debugging)",
    "SourceUsername": "ADO username of the copy source",
    "SourcePassword": "ADO personal access token of the copy source",
    "SourceOrganization": "ADO organization name of the copy source",
    "SourceProject": "Name of the ADO project from which you are copying",
    "SourceTeam": "ADO team name of the copy source",
    "DestinationUsername": "Destination ADO username",
    "DestinationPassword": "Destination ADO personal access token",
    "DestinationOrganization": "Destination ADO organization name",
    "DestinationProject": "Destination ADO project name",
    "DestinationTeam": "Destination ADO team name."
  }
}
```

You can also copy within the same organization.

However, I have not tried copying within the same project as it does not make much sense.

## Mock Azure Storage locally

Mock Azure Storage as it is required for execution.

Press `Ctrl`+`Shift`+`P`, type `Azurite: Start Table Service` and run.

## Run locally

Press `F5` on `index.ts` again.

Go to the Azure Tools extension screen from the Azure icon on the left side of the screen.

Find `CopyProject` in `Local Project` in `WORKSPACE` and right click on it.

Click `Execute Functions Now`. The request body is not loaded, so it can be anything you want.

It should now work.

## To start without debugger

As project root, hit the following command.

```shell-session
yarn start
```

Hit the endpoint in a separate window.

```shell-session
curl http://localhost:7071/api/CopyProject
```

This may be easier than using the debugger.

## If it works

Once it works, prepare Azure Functions and Azure Tables and deploy to production.

## Code Summary

Read the source project and write the destination project via ADO's REST API.

### Realization of difference update

You could delete the entire destination project and then copy all the items again, but that would take too much time.

Therefore, we will use incremental update from the second time onward.

However, it would be a real hassle if we had to use the API to get and check all the contents for a diff update.

Therefore, to detect differences, the ID and the date and time of the last update are stored.

The IDs can be retrieved in a single request, so we can look at the update date and time to determine if an update is necessary.

The problem is that the work item IDs are not retained with the copy.

For example, an item with the number 1500 at the source may have the number 2458 at the destination.

This is an organization-specific ID that cannot be matched.

Therefore, it is necessary to keep a table of the correspondence between the source and destination IDs.

### Filling in missing information

Since a specific individual's PAT is used for execution, there will be missing information.

First, when copying, the submitter of work items and comments becomes the executor of the script.

Also, information on accounts that exist only in the source organization will be lost.

There is no way around this, so the solution is to insert a table with the information in the comments and details fields.

### How to use Azure Tables

We use Azure Tables as a key value store, but accessing it every time can be slow.

Also, it is too much of a hassle in terms of code to hit the asynchronous API frequently.

Therefore, we fetch the contents of a Function at the start of execution and put them in a HashMap during execution.

This HashMap is named `Memo` class.

Then, when the copy is completed, the contents of the HashMap are written back to the Azure Tables.

In this way, key-value store operations can be performed in-memory and handled as a synchronous API.

### Singleton

Many classes are singletons because we don't want to pay the cost of initializing every function call.

The synchronous constructor is a static getter named `instance`.

Asynchronous constructors are static methods named `build()`.

### ESLint & Prettier

At a minimum, ESLint and Prettier code is passed.
