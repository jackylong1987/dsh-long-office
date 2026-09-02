window.__ModuleLoader__.load({
  id: 'dsh-office-reader',
  factory: (require) => {
    const module = { exports: {} }
    const exports = module.exports
    Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' })
    // dsh-office-reader 主要提供服务端 Office 读取/生成工具与路由，此客户端仅为占位(无前端 UI)。
    function apply() {}
    const inject = []
    exports.apply = apply
    exports.inject = inject
    return module.exports
  }
})
