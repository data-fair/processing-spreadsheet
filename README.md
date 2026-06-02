# <img alt="Data FAIR logo" src="https://cdn.jsdelivr.net/gh/data-fair/data-fair@master/ui/public/assets/logo.svg" width="30"> @data-fair/processing-spreadsheet

A plugin that allows the creation and management of datasets from Excel files or zip files containing them.

## Features

- **List sheets** — Lists the sheets and corresponding information in your file
- **Dataset management** — Create or update a file dataset from the desired sheets, configurable from the processing parameters.
- **Graceful stop** — honours the stop signal from the platform and exits cleanly mid-run; optionally, the stop can be ignored to test forced termination after timeout.

## Configuration

The plugin configurations change depending on the datasetMode to be applied. There are three of them:
- `list` to list the different sheets
- `create` to create a new dataset
- `update` to target an existing one

Overall, only the URL in the settings tab remains common, representing a stable URL from which the data file is downloaded (this is the only possible option, there is no repository).

### list

For this mode, you only need to enter the URL in the Parameters tab.

### create

| Tab | Field | Description |
| --- | ----- | ----------- |
| File sheets | `addAllSheets` | Allows you to build datasets directly for all sheets in the file |
| File sheets - Sheets | `add` | Allows you to build a dataset with the corresponding sheet by checking the box |
| File sheets - Sheets | `title` | This corresponds to the name you want for your dataset. By default, the title will be the sheet name. |

### update

| Tab | Field | Description |
| --- | ----- | ----------- |
| Datasets | `datasets` | List of datasets to be updated, taking into account the sheet number and the schema update forcing |
| Datasets - Datasets to update | `dataset` | Name of the dataset to update, selectable from the list of available datasets |
| Datasets - Datasets to update | `sheet` | Name of the sheet from which the update should be made, selectable from the list of available sheets. |
| Datasets - Datasets to update | `forceUpdate` | Indicates whether the scheme update should also be forced |

## Warning

Be careful when fixing dependencies. The project uses ExcelJS **version 4.4.0**. The `npm audit fix` command may downgrade this library, which will definitely crash the processing.

## Release

Processing plugins are fetched from the npm registry with a filter on keyword "data-fair-processings-plugin". So publishing a plugin is as simple as publishing the npm package:

```bash
npm version minor
npm publish
git push --follow-tags
```