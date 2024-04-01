export default class Context {
  static with(ctx: Excel.RequestContext) {
    return {
      sync: this.sync.bind(this, ctx),
      perform: this.perform.bind(this, ctx)
    }
  }

  static async sync<T>(context: Excel.RequestContext, callback: (ctx: Excel.RequestContext) => Promise<T>): Promise<T>
  static async sync<T>(callback: (ctx: Excel.RequestContext) => Promise<T>): Promise<T>
  static async sync<T>(callback_or_context, callback?): Promise<T> {
    callback = callback_or_context instanceof Excel.RequestContext ? callback : callback_or_context

    const batch = async (ctx: Excel.RequestContext) => {
      let result: T = await callback(ctx)

      // TODO: don't load if already loaded
      if (result instanceof Array) {
        result = <T>result.map(item => {
          if (typeof item.load == "function") {
            return item.load()
          } else {
            Object.keys(item).forEach(key => {
              item[key] = item[key].load()
            })
          }
        })
      } else {
        Object.keys(result).forEach(key => {
          result[key] = result[key].load()
        })
      }

      await ctx.sync()

      return result
    }

    if (callback_or_context instanceof Excel.RequestContext) {
      return Excel.run(callback_or_context, batch);
    } else {
      return Excel.run(batch);
    }
  }

  static async perform(context: Excel.RequestContext, callback: (ctx: Excel.RequestContext) => Promise<any>): Promise<any>
  static async perform(callback: (ctx: Excel.RequestContext) => Promise<any>): Promise<any>
  static async perform(callback_or_context, callback?): Promise<any> {
    callback = callback_or_context instanceof Excel.RequestContext ? callback : callback_or_context

    const batch = async (ctx: Excel.RequestContext) => {
      const result = await callback(ctx)
      await ctx.sync()
      return result
    }

    if (callback_or_context instanceof Excel.RequestContext) {
      return await Excel.run(callback_or_context, batch);
    } else {
      return await Excel.run(batch);
    }
  }
}