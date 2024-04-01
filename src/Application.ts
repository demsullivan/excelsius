import Controller from './Controller'

type RegisteredControllerMap = {
  [name: string]: typeof Controller
}

type TaskPaneChangeRequestedEventArgs = {
  viewName: string
  values: any
}

export default class Application extends EventTarget {
  private connectEvents: OfficeExtension.EventHandlerResult<Excel.WorksheetActivatedEventArgs>[] =
    []
  private registeredControllers: RegisteredControllerMap = {}
  private bindingMap: { [controller: string]: string } = {}
  private bindingControllers: { [bindingName: string]: Controller } = {}

  public static start(defaultController?: typeof Controller): Application {
    OfficeExtension.config.extendedErrorLogging = true
    const application = new this()
    application.start()

    if (defaultController) {
      application.register('workbook', defaultController)
      application.activate()
    }

    return application
  }

  async start() {
    await Excel.run(async (context: Excel.RequestContext) => {
      const namedItems = context.workbook.names.load()
      const worksheets = context.workbook.worksheets.load(['names'])
      await context.sync()

      this.setupControllerConnections(context, namedItems)

      await Promise.all(worksheets.items.map(async (sheet: Excel.Worksheet) => {
        return this.setupControllerConnections(context, sheet.names)
      }))

      await context.sync()
    })
  }

  public register(name: string, controller: typeof Controller) {
    this.registeredControllers[name] = controller

    if (this.bindingMap[name]) {
      this.bindingControllers[this.bindingMap[name]] = new controller(this, {})
    }

    Excel.run(async (context: Excel.RequestContext) => {
      const activeWorksheet = context.workbook.worksheets.getActiveWorksheet().load('name')
      const controllerItem = context.workbook.worksheets
        .getActiveWorksheet()
        .names.getItemOrNullObject('controller')
      await context.sync()

      if (!controllerItem.isNullObject) {
        this.handleSheetActivated(name, <Excel.WorksheetActivatedEventArgs>{
          worksheetId: activeWorksheet.name,
          type: Excel.EventType.worksheetActivated
        })
      }
    })
  }

  public updateTaskPane(viewName: string, values: any) {
    this.dispatchEvent(
      new CustomEvent<TaskPaneChangeRequestedEventArgs>('taskPaneChangeRequested', {
        detail: { viewName, values }
      })
    )
  }

  private activate() {
    if (this.registeredControllers.workbook) {
      new this.registeredControllers.workbook(this, <Excel.WorkbookActivatedEventArgs>{
        type: 'WorkbookActivated'
      })
    }
  }

  private async setupControllerConnections(
    context: Excel.RequestContext,
    namedItems: Excel.NamedItemCollection
  ) {
    namedItems.items.map(async (item: Excel.NamedItem) => {
      if (item.scope == 'Worksheet') {
        if (item.name == 'controller') {
          const event = item.worksheet.onActivated.add(
            this.handleSheetActivated.bind(this, item.value)
          )

          this.connectEvents.push(event)
        }
      }

      if (item.name.match(/^controller__/) !== null) {
        const controllerName = item.comment

        Office.context.document.bindings.addFromNamedItemAsync(
          item.name,
          Office.BindingType.Matrix,
          { id: item.name },
          (result) => {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
              this.bindingMap[controllerName] = item.name

              Office.select(`bindings#${item.name}`).addHandlerAsync(
                Office.EventType.BindingDataChanged,
                this.handleBindingDataChanged.bind(this, controllerName)
              )

              Office.select(`bindings#${item.name}`).addHandlerAsync(
                Office.EventType.BindingSelectionChanged,
                this.handleBindingSelectionChanged.bind(this, controllerName)
              )
            }
          }
        )
      }
    })
  }

  private async handleSheetActivated(
    controllerName: string,
    event: Excel.WorksheetActivatedEventArgs
  ) {
    if (this.registeredControllers[controllerName]) {
      console.debug(`activating ${controllerName}`);
      new this.registeredControllers[controllerName](this, event)
    }
  }

  private async handleBindingDataChanged(controllerName: string, event: any) {
    if (this.bindingControllers[event.binding.id]) {
      await this.bindingControllers.dataChanged(event)
    }
  }

  private async handleBindingSelectionChanged(controllerName: string, event: any) {
    if (this.bindingControllers[event.binding.id]) {
      await this.bindingControllers.selectionChanged(event)
    }
  }
}
