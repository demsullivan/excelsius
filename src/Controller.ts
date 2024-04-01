import _ from 'lodash'
import type Application from './Application'
import Context from './Context'
import Binding, {BindingTarget, DataChangedEvent, SelectionChangedEvent} from "./Binding";

type StringIndexedObject = {
  [key: string]: any
}

type InternalEventRecord = {
  name: string;
  listener: EventListener
}

export type TargetDefinition = string | { [name: string]: string }

export type TaskPanePropsFunction = () => Object

export type TaskPaneOptions = {
  view: string,
  props: StringIndexedObject | TaskPanePropsFunction;
}

export default class Controller {
  application: Application
  element!: Excel.Worksheet | Excel.Workbook // | Excel.Range
  internalEvents: InternalEventRecord[] = []
  excelEvents: OfficeExtension.EventHandlerResult<any>[] = []

  targets: TargetDefinition[] = []
  values: string[] = []
  events: string[] = []
  bindings: Binding[]
  valueCache = {}

  taskPane: TaskPaneOptions

  constructor(
    application: Application,
    event: Excel.WorksheetActivatedEventArgs | Excel.WorkbookActivatedEventArgs
  ) {
    this.application = application

    Excel.run(async (context: Excel.RequestContext) => {

      const {element, names} = await Context.sync(async ctx => {
        let element

        if (event.type == 'WorkbookActivated') {
          element = context.workbook
        } else if (event.type == 'WorksheetActivated') {
          element = context.workbook.worksheets.getItem(event.worksheetId)
        }

        return {element, names: element.names}
      });

      this.element = element

      this.setupEventListeners()
      // await this.setupTargets()
      await this.setupValues()

      await context.sync()

      await this.connect()

      await context.sync()

      this.updateTaskPane()
    })
  }

  dispatch(eventName: string, details: CustomEventInit = { detail: {} }) {
    this.application.dispatchEvent(
      new CustomEvent(`excelsius:${eventName}`, details)
    );
  }

  setupEventListeners() {
    if (this.element instanceof Excel.Worksheet) {
      this.excelEvents.push(this.element.onDeactivated.add(this.destroy.bind(this)))
    }

    this.events.forEach((eventName) => {
      if (eventName.match(/^:/)) {
        this.application.addEventListener(`excelsius${eventName}`, this.handleEvent.bind(this, eventName))
        this.internalEvents.push(
          {name: `excelsius${eventName}`, listener: this.handleEvent.bind(this, eventName)}
        )
      } else {
        this.excelEvents.push(
          this.element[`on${eventName}`].add(this.handleEvent.bind(this, eventName))
        )
      }
    })
  }

  async setupTargets() {
    const namedItemTargets = <string[]>this.targets.filter(target => typeof target == "string")
    const rangeTargets = this.targets.filter(target => typeof target != "string")

    const namesAndTables = await Context.sync(async ctx => {
      return [
        ...namedItemTargets.map((target: string) => this.element.names.getItemOrNullObject(target)),
        ...namedItemTargets.map((target: string) => this.element.tables.getItemOrNullObject(target))
      ]
    });

    namesAndTables.forEach((name_or_table: Excel.NamedItem | Excel.Table) => {
      if (name_or_table.isNullObject) return;

      if (name_or_table instanceof Excel.NamedItem) {
        this.createBinding(name_or_table.getRange(), name_or_table.name)
      } else if (name_or_table instanceof Excel.Table) {
        this.createBinding(name_or_table, name_or_table.name)
      }
    });

    if (this.element instanceof Excel.Worksheet) {
      const element = <Excel.Worksheet>this.element

      rangeTargets.forEach((target: TargetDefinition) => {
        Object.keys(target).forEach(name => {
          this.createBinding(element.getRange(target[name]), name)
        })
      })
    }
  }

  createBinding(target: BindingTarget, name: string) {
    const binding = new Binding(target, name)

    const dataChangedHandler = this[`${name}TargetDataChanged`]
    const selectionChangedHandler = this[`${name}TargetSelectionChanged`]

    if (dataChangedHandler !== undefined) {
      binding.addEventListener('dataChanged', dataChangedHandler.bind(this))
    }

    if (selectionChangedHandler !== undefined) {
      binding.addEventListener('selectionChanged', selectionChangedHandler.bind(this))
    }

    this[`${name}Target`] = binding

    this.bindings.push(binding);
  }

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
      this.excelEvents.forEach(async (event) => {
        await Excel.run(event.context, async (context) => {
          event.remove()
          await context.sync()
        })
      })

      this.internalEvents.forEach((eventRecord: InternalEventRecord) => {
        this.application.removeEventListener(eventRecord.name, eventRecord.listener)
      });

      this.bindings.forEach(async binding => await binding.destroy())

      this.excelEvents = []
      this.internalEvents = []
      this.bindings = []

      resolve()
    })
  }

  async handleEvent(eventName: string, event: any) {
    const handler = this[`handle${eventName.replace(/^:/, "")}`]

    if (handler) {
      handler.bind(this)(event)
    }

    this.updateTaskPane()
  }

  updateTaskPane() {
    let props

    if (this.taskPane === undefined) return;

    if (typeof this.taskPane.props == "function") {
      props = this.taskPane.props.call(this)
    } else {
      props = this.taskPane.props
    }

    this.application.updateTaskPane(this.taskPane.view, props)
  }

  async connect() {
  }

  async disconnect() {
  }

  async fetch<T>(callback: (context: Excel.RequestContext) => any): Promise<T> {
    return await Excel.run(async (context: Excel.RequestContext) => {
      const result = callback(context)
      await context.sync()
      return result
    })
  }
}
