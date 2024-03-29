# Excelsius
A lightweight controller framework for Office.js Excel Add-ins, inspired by Stimulus.js

> [!CAUTION]
> This library has not been released yet and is still under active development.
> If you're interested in using it, please drop me a note [here](https://github.com/demsullivan/excelsius/issues/1)

## Install

Install via your favourite package manager. For example:

```
npm install --save excelsius
```

## Getting Started

Spin up an Office.js Excel Add-in, following the instructions [here](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator)

Initialize Excelius in your Add-in Javascript or Typescript code. For example:

```typescript
import { Application } from 'excelsius'

window.Excelsius = Application.start()
```

## Building A Simple Controller

Excelsius Controllers can be attached to workbooks, worksheets, and named ranges. Here's a simple example:

```typescript
// my-add-in/src/controllers/HelloWorldController.ts
import { Controller } from 'excelsius'

export default class extends Controller {
  declare element: Excel.Worksheet

  async connect() {
    this.element.getRange("A1").value = "Hello world!"
  }
}

// my-add-in/src/index.ts
import { Application } from 'excelsius'
import HelloWorldController from './controllers/HelloWorldController'

window.Excelsius = Application.start()
window.Excelsius.register('hello_world', HelloWorldController)
```

Then, to connect a controller to a specific Worksheet in your Workbook:

- Click on the "Formulas" tab
- Click the "Name Manager" button
- Click "New" to create a new name
- Set the following properties on the name:
  - Name: controller
  - Scope: Choose one of the Worksheets in your document
  - Refers to: `="hello_world"`

Once this is done, open (or refresh) your Add-in in Excel, and you should see "Hello world!" printed in cell A1 of the worksheet you selected for the name.

### Attaching a Controller to a Workbook

### Attaching a Controller to a Named Range

## Interacting with the Taskpane

## Events

## Values

## Bindings

