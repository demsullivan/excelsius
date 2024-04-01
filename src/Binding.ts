export type BindingTarget = Excel.Table | Excel.Range

export type DataChangedEventArgs = {
  binding: Binding
}

export type SelectionChangedEventArgs = {
  binding: Binding
}

export class DataChangedEvent extends CustomEvent<DataChangedEventArgs> {
  constructor(detail: CustomEventInit<DataChangedEventArgs>) {
    super("dataChanged", detail);
  }
}

export class SelectionChangedEvent extends CustomEvent<SelectionChangedEventArgs> {
  constructor(detail: CustomEventInit<SelectionChangedEventArgs>) {
    super("selectionChanged", detail);
  }
}

export default class Binding extends EventTarget {
  name: string
  bindingType: Excel.BindingType
  binding: Excel.Binding

  constructor(target: BindingTarget, name: string) {
    super()

    let itemName
    this.bindingType = target instanceof Excel.Table ? Excel.BindingType.table : Excel.BindingType.range

    this.name = name

    if (target instanceof Excel.Table) {
      itemName = target.name
    } else if (target instanceof Excel.Range) {
      itemName = target.address
    } else {
      itemName = target
    }

    Excel.run(async (context: Excel.RequestContext) => {
      if (target instanceof Excel.Range) {
        this.binding = context.workbook.bindings.add(target, this.bindingType, this.name)
      } else {
        this.binding = context.workbook.bindings.addFromNamedItem(itemName, this.bindingType, this.name);
      }

      this.binding.onDataChanged.add(this.handleBindingDataChanged.bind(this));
      this.binding.onSelectionChanged.add(this.handleBindingSelectionChanged.bind(this));

      await context.sync();
    });

  }

  handleBindingDataChanged(event: Excel.BindingDataChangedEventArgs) {
    this.dispatchEvent(new DataChangedEvent({detail: {binding: this}}))
  }

  handleBindingSelectionChanged(event: Excel.BindingSelectionChangedEventArgs) {
    this.dispatchEvent(new SelectionChangedEvent({detail: {binding: this}}))
  }

  async destroy() {
    return await Excel.run(this.binding.context, async ctx => {
      this.binding.delete()
      await ctx.sync()
    });
  }
}