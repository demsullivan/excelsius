import _ from 'lodash'
import type Application from './Application'

type StringIndexedObject = {
  [key: string]: any
}

export default class Controller {
  application: Application
  element!: Excel.Worksheet | Excel.Workbook
  excelEvents: OfficeExtension.EventHandlerResult<any>[] = []

  events: string[] = []
  bindings: string[] = []
  values: string[] = []

  valueCache = {}

  constructor(
    application: Application,
    event: Excel.WorksheetActivatedEventArgs | Excel.WorkbookActivatedEventArgs
  ) {
    this.application = application

    Excel.run(async (context: Excel.RequestContext) => {
      if (event.type == 'WorkbookActivated') {
        this.element = context.workbook.load()
      } else if (event.type == 'WorksheetActivated') {
        this.element = context.workbook.worksheets.getItem(event.worksheetId).load()
      }

      this.setupEventListeners()
      this.setupBindings()
      await this.setupValues()

      await context.sync()

      await this.connect()

      await context.sync()
    })
  }

  setupEventListeners() {
    if (this.element instanceof Excel.Worksheet) {
      this.excelEvents.push(this.element.onDeactivated.add(this.destroy.bind(this)))
    }

    this.events.forEach((eventName) => {
      this.excelEvents.push(
        this.element[`on${eventName}`].add(this.handleEvent.bind(this, eventName))
      )
    })
  }

  setupBindings() {}

  async setupValues() {
    this.element.names.load()
    await this.element.context.sync()

    const namedItems: StringIndexedObject = this.values.map((value: string) => {
      const namedItemName = `value__${_.snakeCase(value)}`

      this[`${value}Value`] = this.getValue.bind(this, namedItemName)
      this[`set${value}Value`] = this.setValue.bind(this, namedItemName)
      return this.element.names.getItemOrNullObject(namedItemName).load()
    })

    await this.element.context.sync()

    this.valueCache = namedItems.reduce((valueMap: StringIndexedObject, item: Excel.NamedItem) => {
      valueMap[item.name] = item.isNullObject ? null : item.value
      return valueMap
    }, {})
  }

  getValue(name: string) {
    return this.valueCache[name]
    // if (this.valueCache[name]) return this.valueCache[name]

    // const namedItem = this.element.names.getItemOrNullObject(name).load()
    // await this.element.context.sync()

    // if (namedItem.isNullObject) {
    //   return null
    // } else {
    //   return namedItem.value
    // }
  }

  async setValue(name: string, value: any) {
    // TODO: support setting values on a range
    let namedItem = null
    try {
      namedItem = this.element.names.getItem(name).load()
      await this.element.context.sync()
      namedItem.delete()
    } catch (e) {
      namedItem = null
    } finally {
      this.element.names.add(name, `="${value}"`, namedItem?.comment)
      this.valueCache[name] = value
      await this.element.context.sync()
    }
  }

  async destroy(event: Excel.WorksheetDeactivatedEventArgs) {
    await this.disconnect()

    await new Promise<void>((resolve) => {
      this.excelEvents.forEach((event) => {
        Excel.run(event.context, async (context) => {
          event.remove()
          await context.sync()
        })
      })

      this.excelEvents = []

      resolve()
    })
  }

  async handleEvent(eventName: string, event: any) {
    const handler = this[`handle${eventName}`]

    if (handler) {
      handler.bind(this)(event)
    }
  }

  updateTaskPane(viewName: string, values: any) {
    this.application.updateTaskPane(viewName, values)
  }

  async connect() {}

  async disconnect() {}

  async fetch<T>(callback: (context: Excel.RequestContext) => any): Promise<T> {
    return await Excel.run(async (context: Excel.RequestContext) => {
      const result = callback(context)
      await context.sync()
      return result
    })
  }
}
