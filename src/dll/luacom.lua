if not luacom then
  local init, err1, err2 = loadlib("luacom-1.3.dll","luacom_openlib")
  assert (init, (err1 or '')..(err2 or ''))
  init()
end

return nil
