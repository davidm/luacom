-- Testes luacom





teste_progid = "TESTECTL.TestectlCtrl.1"





Unk = luacom.GetIUnknown





-- Funcoes auxiliares

function table2string(table)



  if type(table) ~= "table" then

    return tostring(table)

  end



  local i, v = next(table, nil)



  local s = "{"

  local first = 1



  while i ~= nil do

    

    if first ~= 1 then

      s = s..", "            

    else

      first = 0

    end



    s = s..table2string(v)

    i,v = next(table, i)    

  end



  s = s.."}"



  return s



end





function number_test(start)

  if number_test_n == nil then

    number_test_n = 0

  end



  if type(start) == "number" then

    number_test_n = start

  else

    number_test_n = number_test_n + 1

  end



  if type(start) == "string" then

	print("Teste "..number_test_n..": "..start);

  else

	print("Teste "..number_test_n);

  end

end



nt = number_test





-- Testa tratamento de entradas incorretas

function test_wrong_parameters()

  print("\n=======> test_wrong_parameters")

  nt(1)

  res = call(luacom.CreateObject,{{table = "aa"}},"x")

  assert(res == nil)



  nt()

  obj = call(luacom.CreateObject,{},"x")

  assert(obj == nil)



  nt()

  obj = call(luacom.ImplInterface, {}, "x")

  assert(obj == nil)



  nt()

  obj = call(luacom.ImplInterface,{{}, "blablabla", "blablabla"},"x")

  assert(obj == nil)



  nt()

  obj = call(luacom.ImplInterface, {{}, "InetCtls.Inet", "blablabla"}, "x")

  assert(obj == nil)



  -- ProgID's inexistentes

  nt()

  obj = luacom.CreateObject("blablabla")

  assert(obj == nil)



  nt()

  obj = luacom.CreateObject(1)

  assert(obj == nil)



end









-- Teste simples

function test_simple()

  print("\n=======> test_simple")

  nt(1)



  obj = luacom.CreateObject("MSComCtl2.Animation")

  assert(obj ~= nil)



  nt()

  assert(type(obj.AboutBox) == "function")



  nt()

  assert(type(obj.BackColor) == "number")



  nt()

  obj.BackColor = 2

  assert(obj.BackColor == 2)



  nt()

  table = {}



  table.TestArray1 = function(self, array)

    assert(table2string(array) == table2string({"1","2"}))

  end



  obj = luacom.ImplInterface(

    table, "TESTE.Teste", "IDataConversionTest" )



  assert(obj)



  obj:TestArray1({1,2})

end









-- Teste de Implementacao de Interface via tabelas

function test_iface_implementation()



  print("\n=======> test_iface_implementation")



  nt(1)



  iface_ok = 0



  iface = {}



  iface.TestEvent = function(self)

    iface_ok = iface_ok + 1

  end



  iface.TestManyParams = function(self,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,p13,p14,p15)

    assert(p1 == 1)

    assert(p2 == 2)

    assert(p3 == 3)

    assert(p4 == 4)

    assert(p5 == 5)

    assert(p6 == 6)

    assert(p7 == 7)

    assert(p8 == 8)

    assert(p9 == 9)

    assert(p10 == 10)

    assert(p11 == 11)

    assert(p12 == 12)

    assert(p13 == 13)

    assert(p14 == 14)

    assert(p15 == 15)

  end



  obj = luacom.ImplInterface(iface, teste_progid, "_DTestectlEvents")

  assert(obj)



  nt()

  obj:TestEvent()

  assert(iface_ok == 1) 



  nt()

  obj:TestManyParams(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)



  obj = nil

  collectgarbage()



end



function test_luadisp_in_params()

  print("\n=======> test_luadisp_in_params")

  nt(1)



  local iface = {}



  iface.TestParam1 = function(self, p1, p2, p3)

    return p1

  end



  iface.TestParam2 = function(self, p1, p2, p3)

    return p2

  end





  iface.TestParam3 = function(self, p1, p2, p3)

    return p3

  end



  obj = luacom.ImplInterface(iface, teste_progid, "_DTestectl")

  assert(obj)





  -- Testa ordem dos parametros



  local res = obj:TestParam1(1,2,3)

  assert(res == 1)



  res = obj:TestParam2(1,2,3)

  assert(res == 2)



  res = obj:TestParam3(1,2,3)

  assert(res == 3)

end







function test_luadisp_inout_params()

  print("\n=======> test_luadisp_inout_params")



  nt(1)



  local iface = {}



  iface.TestInOutshort = function(self, p1)

    return p1*2

  end



  iface.TestInOutfloat = function(self, p1)

    return p1*2

  end



  iface.TestInOutlong = function(self, p1)

    return p1*2

  end



  iface.TestInOutdouble = function(self, p1)

    return p1*2

  end





  obj = luacom.ImplInterface(iface, teste_progid, "_DTestectl")

  assert(obj)



  nt()



  local res = call(obj.TestInOutshort, {obj,10}, "x")

  assert(res == 20)



  nt()



  local res = call(obj.TestInOutfloat, {obj,2.5}, "x")

  assert(res == 5)



  nt()

  local res = call(obj.TestInOutlong, {obj,100000}, "x")

  assert(res == 200000)



  nt()



  local res = call(obj.TestInOutdouble, {obj,2.5}, "x")

  assert(res == 5)



end





function test_luadisp_out_params()

  print("\n=======> test_luadisp_out_params")



  nt()



  local iface = {}



  iface.TestOutParam = function(self)

    return 1000

  end



  iface.Test2OutParams = function(self)

    return 1.1, tostring(2*1.1)

  end



  iface.TestRetInOutParams = function(self, p1, p2)

    return p1*p2, p1, p2

  end



  obj = luacom.ImplInterface(iface, teste_progid, "_DTestectl")

  assert(obj)



  nt()



  local res = call(obj.TestOutParam, {obj}, "x")

  assert(res == 1000)



  nt()



  local res1, res2 = call(obj.Test2OutParams, {obj}, "x")

  assert(res1 == 1.1 and tonumber(res2) == 2*res1)



  nt()

  local res1, res2, res3 = call(obj.TestRetInOutParams, {obj,2,3}, "x")

  assert(res1 == 6)

  assert(res2 == 2)

  assert(res3 == 3)



end



function test_luadisp_propget()

  print("\n=======> test_luadisp_propget")



  nt(1)



  local iface = {}



  iface.TestPropParam = {{1, "a"},{"b", 2}}





  obj = luacom.ImplInterfaceFromTypelib(iface, 

        "c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )

  

  assert(obj)



  nt()

  assert(obj.TestPropParam[1][1] == 1)



  nt()

  assert(obj.TestPropParam[1][2] ==  "a")



  nt()

  assert(obj.TestPropParam[2][1] == "b")



  nt()

  assert(obj.TestPropParam[2][2] == 2)



  

end



function test_luadisp_propput()

  print("\n=======> test_luadisp_propput")



  nt(1)



  local iface = {}



  iface.TesteProp = {}





  obj = luacom.ImplInterfaceFromTypelib(iface, 

        "c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )

  

  assert(obj)



  -- propput normal

  nt()

  obj.Teste = 100

  assert(obj.Teste == 100)



  -- Propput parametrizado

  nt()

  obj:setTesteProp(1,1)

  assert(obj:TesteProp(1) == 1)



  nt()

  obj:setTesteProp(10,11)

  assert(obj:TesteProp(10) == 11)



  nt()

  obj:setTesteProp(20,21)

  assert(obj:TesteProp(20) == 21)

end









-- Teste de stress

function test_stress()

  print("\n=======> test_stress")

  nt(1)



  iface = {}

  iface.BeforeNavigate = function (self)

    iface_ok = iface_ok + 1

  end



  obj = luacom.ImplInterface(iface, "InternetExplorer.Application", "DWebBrowserEvents")

  assert(obj)



  i = 1

  iface_ok = 1



  while i < 100000 do

    obj:BeforeNavigate()

    i = i + 1

  end



  nt()

  assert(i == iface_ok)



  obj = nil

  collectgarbage()

end















-- Teste de Connection Points

function test_connection_points()

  print("\n=======> test_connection_points")

  nt(1)



  obj = luacom.CreateObject("MSComCtl2.UpDown")

  assert(obj)



  nt()

  events = {}



  events_ok = nil



  events.Change = function(self)

    events_ok = 1

  end



  evt = luacom.ImplInterface(events, "MSComCtl2.UpDown", "DUpDownEvents")

  assert(evt)



  nt()

  luacom.addConnection(obj, evt)



  nt()


  obj.Value = obj.Value + 1



  assert(events_ok)



  nt()

  luacom.releaseConnection(obj)



  events_ok = nil



  nt()

  luacom.Connect(obj, events)



  nt()

  obj.Value = obj.Value + 1



  assert(events_ok)



  nt()

  luacom.releaseConnection(obj)



  obj = nil

  events = nil

  collectgarbage()



end









-- Teste de propget

function test_propget()

  print("\n=======> test_propget")



  nt(1)



  local iface = {}

  iface.Value = 1



  local evt = luacom.ImplInterface(iface, "MsComCtl2.UpDown", "IUpDown")

  assert(evt)



  assert(evt.Value == iface.Value)



  --nt()

  --obj = luacom.CreateObject("COMCTL.TabStrip.1")

  --t = obj.Tabs

  

end





-- Teste de propput

function test_propput()

  print("\n=======> test_propput")



  nt(1) -- propput simples



  obj = luacom.CreateObject("MSCAL.Calendar.7")

  assert(obj)



  -- Propput normal

  obj.Day = 1

  assert(obj.Day == 1)



  nt() -- propput parametrizado



  obj = luacom.CreateObject("MSFlexGridLib.MSFlexGrid")

  assert(obj)



  obj.Cols = 5

  obj.Rows = 5



  obj:setTextMatrix(1,1,"abc")

  assert(obj:TextMatrix(1,1) == "abc")



  nt()

  obj:setTextMatrix(3, 2, 9)

  assert(tonumber(obj:TextMatrix(3,2)) == 9)

end









function test_index_fb()

  print("\n=======> test_index_fb")

  nt(1)



  local obj = luacom.CreateObject("MSComCtl2.Animation")

  assert(obj)



  local tm = gettagmethod(tag(obj), "index")

  assert(tm)



  nt()

  assert(call(tm, {nil,nil}, "x") == nil)

  

  nt()

  assert(call(tm,{{}, nil}, "x") == nil)



  nt()

  assert(call(tm,{{}, "test"}, "x") == nil)



  nt()

  assert(call(tm,{obj, nil}, "x") == nil)



  nt()

  local f = tm(obj, "blublublabla")

  assert(f == nil)



  nt()

  local f = tm(obj, "AboutBox")

  assert(type(f) == "function")

end









function test_method_call_without_self()

  print("\n=======> test_method_call_without_self")

  nt(1)



  local obj = luacom.CreateObject("MSComCtl2.Animation")

  assert(obj)



  obj:AboutBox()



  nt()

  obj.AboutBox(obj)



  nt()

  obj.AboutBox({})



  nt()

  local res = call(obj.AboutBox, {1}, "x")

  assert(res == nil)



  nt()

  res = call(obj.AboutBox, {}, "x")

  assert(res == nil)





end



function test_in_params()

  print("\n=======> test_in_params")



  obj = luacom.CreateObject(teste_progid)

  assert(obj)



  -- Testa ordem dos parametros



  nt(1)

  local res = obj:TestParam1(1,2,3)

  assert(res == 1)



  res = obj:TestParam2(1,2,3)

  assert(res == 2)



  res = obj:TestParam3(1,2,3)

  assert(res == 3)



end



function test_inout_params()

  print("\n=======> test_inout_params")



  iupobj = iupolecontrol{progid = "TESTECTL.TesteCtrl.1"}

  assert(iupobj)

  obj = iupobj.LUACOM



  nt(1)



  local res = call(obj.TestInOutshort, {obj,10}, "x")

  assert(res ~= nil)

  assert(res[1] == 20)



  nt()



  local res = call(obj.TestInOutfloat, {obj,2.5}, "x")

  assert(res ~= nil)

  assert(res[1] == 5)



  nt()

  local res = call(obj.TestInOutlong, {obj,100000}, "x")

  assert(res ~= nil)

  assert(res[1] == 200000)



end



----------------------------------------------------------------------

----------------------------------------------------------------------



function test_out_params()

  print("\n=======> test_out_params")



  iupobj = iupolecontrol{progid = "TESTECTL.TesteCtrl.1"}

  assert(iupobj)

  obj = iupobj.LUACOM



  nt(1)



  local res = call(obj.TestOutParam, {obj}, "x")

  assert(res ~= nil and res[1] == 1000)



  nt()



  local res = call(obj.Test2OutParams, {obj}, "x")

  assert(res ~= nil)

  assert(res[1] == 1.1 and tonumber(res[2]) == res[1])



  nt()

  local res = call(obj.TestRetInOutParams, {obj,2,3}, "x")

  assert(res ~= nil)

  assert(res[1] == 6)

  assert(res[2] == 2)

  assert(res[3] == 3)



end



----------------------------------------------------------------------

----------------------------------------------------------------------

-- teste de safe array

----------------------------------------------------------------------



function test_safearrays()

  print("\n=======> test_safearrays")



  teste = {}



  nt(1)



  teste.Teste = 1



  obj = luacom.ImplInterfaceFromTypelib(

    teste, "c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )



  assert(obj)



  array = {}



  local i=1

  local j=1

  local k=1



  while i <= 3 do



    if array[i] == nil then

      array[i] = {}

    end

    j = 1

    while j <= 3 do



      if array[i][j] == nil then

        array[i][j] = {}

      end

      k = 1

      while k <= 3 do

        array[i][j][k] = i*9+j*3+k

        k = k + 1

      end



      j = j + 1

    end



    i = i + 1

  end



  teste.TestSafeArray = function(self, array_param)

      assert(table2string(array_param) == table2string(%array))

    end



  local i = 0

  while i < 10 do

    obj:TestSafeArray(array)

    i = i + 1

  end



  nt()



  -- testa tratamento de erro na conversao



  teste.TestSafeArray2 = function(self, array_param)

	  print(table2string(array_param))

	  return 1

    end



  teste.TestSafeArrayVariant = function(self, array_param)

    return array_param

  end



  local res = call(obj.TestSafeArray2,{obj,{}}, "x")

  assert(res == nil)



  local res = call(obj.TestSafeArray2,{obj,{1,2,{3}}}, "x")

  assert(res == nil)



  local res = call(obj.TestSafeArray2,{obj,{{1},{2},{3,4}}}, "x")

  assert(res == nil)



  local res = call(obj.TestSafeArray2,{obj,{{{1,2}},{{2,1}},{{3}}}}, "x")

  assert(res == nil)



  -- steals a userdata...

  userdata = obj["_USERDATA_REF_"]

  local res = call(obj.TestSafeArrayVariant,{obj,{{userdata, userdata}}},"x")

  assert(res == nil)



  nt()  -- tests long safe arrays



  local i = 0

  array = {}



  for i=1,100 do

    array[i] = i

  end



  a = obj:TestSafeArrayVariant(array)



  for i=1,100 do

    assert(a[i] == i)

  end



  array_test = {{1},{"1"},{1},{1},{1},{1},{1},{1},{1},{1},{1},{1},{1}}

  a = obj:TestSafeArrayVariant(array_test)

  assert(table2string(a) == table2string(array_test))



  nt()



  teste.Teste = 10

  teste.TestSafeArrayIDispatch = function(self, array)

    local i = 1

    

    while i < 10 do

      print(tostring(i).." => "..tostring(array[i].Teste))

      i = i + 1

    end

  end



  obj = luacom.ImplInterfaceFromTypelib(

    teste,"c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )



  assert(obj)



  array = {}



  local i = 1



  while i < 10 do

    array[i] = luacom.ImplInterfaceFromTypelib(

      teste, "c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )



    i = i + 1

  end



  obj:TestSafeArrayIDispatch(array)



  collectgarbage()



end







--

-- Teste de luacom.NewObject

--



function test_NewObject()

  print("\n=======> test_NewObject")

  nt(1)



  local tm_gettable = 

    function(table, index)

	  rawset(table, "Accessed", 1)

	  return rawget(table, index)

    end



  local tm_settable = 

    function(table, index, value)

	  rawset(table, index, value)

    end



  impl = {}



  testtag = newtag()

  settag(impl, testtag)



  settagmethod(testtag, "gettable", tm_gettable)

  settagmethod(testtag, "settable", tm_settable)



  obj = luacom.NewObject(impl, "MsComCtl2.UpDown")

  assert(obj)



  obj.Value = 1

  assert(obj.Value == 1 and impl.Accessed == 1)

end





--

-- Teste da conversao de dados

--



function test_DataTypes()

  print("\n=======> test_DataTypes")



  teste = {}



  teste.Test = function(self, in_param, in_out_param, out_param)

    return in_out_param, in_param, in_param

  end



  obj = luacom.ImplInterface(teste, "TESTE.Teste", "IDataConversionTest")

  assert(obj)



  nt(1)  -- Datas



  teste.TestDATE = teste.Test





  -- this one should work  

  date = "29/2/1996 10:00:00"

  date2 = "1/1/2001 01:00:00"

  date_res1, date_res2, date_res3 = obj:TestDATE(date, date2)

  assert(date_res1 == date2)

  assert(date_res2 == date)

  assert(date_res3 == date)


  -- this ones should fail

  date = "29/2/1997 10:00:00"

  res = call(obj.TestDATE,{obj, date, date},"x")

  assert(res == nil)



  date = "1/1/2001 24:00:00"

  res = call(obj.TestDATE,{obj, date, date},"x")

  assert(res == nil)







  nt("Currency") -- teste de currency



  teste.TestCURRENCY = teste.Test



  cy = "R$ 1.000.001,15"

  cy2 = 988670.12

  cy_res2, cy_res, cy_res3 = obj:TestCURRENCY(cy, cy2)

  assert(cy_res == 1000001.15)

  assert(cy_res2 == cy2)

  assert(cy_res3 == 1000001.15)



  cy = 12345.56

  cy_res = obj:TestCURRENCY(cy, cy)

  assert(cy_res == cy)



  -- Este deve falhar



  cy = "R$ 1,000,001.15"

  res = call(obj.TestCURRENCY,{obj, cy},"x")

  assert(res == nil)



  cy = {}

  res = call(obj.TestCURRENCY,{obj, cy},"x")

  assert(res == nil)



  nt() -- teste de booleanos



  teste.TestBool = teste.Test



  b = 1

  b2 = 0

  b_res2, b_res, b_res3 = obj:TestBool(b, b2)

  assert(b == b_res)

  assert(b2 == b_res2)

  assert(b_res3 == b)



  b = 0

  b_res = obj:TestBool(b, b)

  assert(b == b_res)



  nt() -- teste de unsigned char



  teste.TestChar = teste.Test

  

  uc = 100

  uc2 = 150

  uc_res2, uc_res, uc_res3 = obj:TestChar(uc, uc2)

  assert(uc_res == uc)

  assert(uc_res2 == uc2)

  assert(uc_res3 == uc)



  -- este deve falhar

  uc = 1000

  res = call(obj.TestChar, {obj, uc, uc}, "x")

  assert(res == nil)





  -- teste de VARIANT



  teste.TestVARIANT = teste.Test

  

  nt("VARIANT - number")

  v = 100

  v2 = 150

  v_res2, v_res, v_res3 = obj:TestVARIANT(v, v2)

  assert(v_res == v)

  assert(v_res2 == v2)

  assert(v_res3 == v)



  nt("VARIANT - String")

  v = "abcdef"

  v2 = "ghijog"

  v_res2, v_res, v_res3 = obj:TestVARIANT(v, v2)

  assert(v_res == v)

  assert(v_res2 == v2)

  assert(v_res3 == v)



  nt("VARIANT - mixed")

  v = "abcdef"

  v2 = 123

  v_res2, v_res, v_res3 = obj:TestVARIANT(v, v2)

  assert(v_res == v)

  assert(v_res2 == v2)

  assert(v_res3 == v)



  nt("VARIANT - IDispatch")

  v = obj

  v2 = obj

  v_res2, v_res, v_res3 = obj:TestVARIANT(v, v2)

  assert(luacom.GetIUnknown(v_res)== luacom.GetIUnknown(v))

  assert(luacom.GetIUnknown(v_res2)== luacom.GetIUnknown(v2))

  assert(luacom.GetIUnknown(v_res3)== luacom.GetIUnknown(v))



  nt("VARIANT - IUnknown")

  v = luacom.GetIUnknown(obj)

  v2 = v

  v_res2, v_res, v_res3 = obj:TestVARIANT(v, v2)

  assert(v_res == v)

  assert(v_res2 == v2)

  assert(v_res3 == v)







  nt("Enum") -- Teste de Enum



  teste.TestENUM = teste.Test



  e = 3

  e2 = 4

  e_res2, e_res, e_res3 = obj:TestENUM(e, e2)

  assert(e_res == e)

  assert(e_res2 == e2)

  assert(e_res3 == e)



  nt("BSTR") -- Teste de BSTR



  teste.TestBSTR = teste.Test



  s = "abdja"

  s2 = "fsjakfhdksjfhkdsa"

  s_res2, s_res, s_res3 = obj:TestBSTR(s, s2)

  assert(s_res == s)

  assert(s_res2 == s2)

  assert(s_res3 == s)



  nt("Short")



  teste.TestShort = teste.Test



  p1 = 10

  p2 = 100

  res2, res1, res3 = obj:TestShort(p1, p2)

  assert(p1 == res1)

  assert(p2 == res2)

  assert(p1 == res3)



  nt("Long")



  teste.TestLong = teste.Test



  p1 = 1234567

  p2 = 9876543

  res2, res1, res3 = obj:TestLong(p1, p2)

  assert(p1 == res1)

  assert(p2 == res2)

  assert(p1 == res3)



  nt("Float")



  teste.TestFloat = teste.Test



  p1 = 10.111

  p2 = 123.8212

  res2, res1, res3 = obj:TestFloat(p1, p2)

  assert(abs(p1-res1) < 0.00001)

  assert(abs(p2-res2) < 0.00001)

  assert(abs(p1-res3) < 0.00001)



  nt("Double")



  teste.TestDouble = teste.Test



  p1 = 9876.78273823

  p2 = 728787.889978

  res2, res1, res3 = obj:TestDouble(p1, p2)

  assert(p1 == res1)

  assert(p2 == res2)

  assert(p1 == res3)



  nt("Int")



  teste.TestInt = teste.Test



  p1 = 1000

  p2 = 1001

  res2, res1, res3 = obj:TestInt(p1, p2)

  assert(p1 == res1)

  assert(p2 == res2)

  assert(p1 == res3)



  nt("IDispatch")



  teste.TestIDispatch = teste.Test



  p1 = obj

  p2 = obj

  res2, res1, res3 = obj:TestIDispatch(p1, p2)

  assert(luacom.GetIUnknown(res1) == luacom.GetIUnknown(obj))

  assert(luacom.GetIUnknown(res2) == luacom.GetIUnknown(obj))

  assert(luacom.GetIUnknown(res3) == luacom.GetIUnknown(obj))



  nt("IUnknown")



  teste.TestIUnknown = teste.Test



  p1 = luacom.GetIUnknown(obj)

  p2 = luacom.GetIUnknown(obj)

  res2, res1, res3 = obj:TestIUnknown(p1, p2)

  assert(p1 == res1)

  assert(p2 == res2)

  assert(p1 == res3)



end



function test_USERDEF_PTR()

  print("\n=======> test_USERDEF_PTR")



  teste_table = {}



  -- por default simplesmente retorna

  teste_table.teste = function(self, value)

    return value

  end



  myindex = function(self)

    return rawget(self, "teste")

  end



  t = newtag()



  settag(teste_table, t)

  settagmethod(t, "index", myindex)



  teste = luacom.ImplInterface(teste_table, "TESTE.Teste", "ITest1")

  teste2 = luacom.ImplInterface(teste_table, "TESTE.Teste", "IStruct1")

  assert(teste)





  -- outra opcao



  teste_disp = function(self, value)

    local t = {}

    o2 = luacom.ImplInterface(t, "TESTE.Teste", "ITest1")

    return o2

  end



  nt(1)



  n = teste:up_userdef_enum(3)

  assert(n == nil)



  nt()

  n = teste:up_ptr_userdef_enum(4)

  assert(n == 4)



  nt()



  teste_table.up_ptr_disp = teste_disp



  n = teste:up_ptr_disp(teste2)

  assert(luacom.GetIUnknown(n) == luacom.GetIUnknown(o2))



  nt()



  n = teste:up_ptr_userdef_disp(teste2)

  assert(n == nil)



  nt()



  teste_table.up_ptr_ptr_userdef_disp = teste_disp



  n = teste:up_ptr_ptr_userdef_disp(teste2)

  assert(luacom.GetIUnknown(n) == luacom.GetIUnknown(o2))



  nt()

  n = teste:up_userdef_alias(1)

  assert(n == nil)



  nt()

  n = teste:up_ptr_userdef_alias(1)

  assert(n == 1)



  n = teste:up_userdef_alias_userdef_alias(1)

  assert(n == nil)



  nt()

  n = teste:up_userdef_disp(teste2)

  assert(n == nil)



  nt()



  teste_table.up_udef_alias_ptr_udef_alias_ptr_disp = teste_disp



  n = teste:up_udef_alias_ptr_udef_alias_ptr_disp(teste2)

  assert(luacom.GetIUnknown(n) == luacom.GetIUnknown(o2))





  nt()

  teste_table.up_userdef_disp = function(self, value)

    o2 = value

  end



  teste:up_userdef_disp(teste2)

  assert(luacom.GetIUnknown(teste2) == luacom.GetIUnknown(o2))



  nt("Unknown & Userdef")



  teste_unk = function(self, value)

    local t = {}

    o2 = luacom.ImplInterface(t, "TESTE.Teste", "ITest1")

    o2 = luacom.GetIUnknown(o2)

    return o2

  end

  

  teste_table.up_udef_alias_ptr_udef_alias_ptr_unk = teste_unk



  n = teste:up_udef_alias_ptr_udef_alias_ptr_unk(luacom.GetIUnknown(teste2))

  assert(n == o2)



  nt("Unknown & Userdef")

  teste_table.up_userdef_unk = function(self, value)

    o2 = value

  end

  

  teste:up_userdef_unk(luacom.GetIUnknown(teste2))

  assert(luacom.GetIUnknown(teste2) == o2)



end



function test_field_redefinition()

  print("\n=======> test_field_redefinition")  

  nt(1)

  local luaobj = {}

  luaobj.getput = 1
  luaobj.get = 2
  luaobj.func = function() return 3 end

  local obj = luacom.ImplInterface(luaobj, "TESTE.Teste", "ITest1")

  assert(obj)

  assert(obj.getput == 1)
  assert(obj.get == 2)
  assert(obj:func() == 3)

  nt()

  obj.getput = function() return 2 end

  assert(obj:getgetput() == 2)

  local settable = function(t, i,v )
    t[i] = v
    return 1
  end


  nt()

  local res = call(settable,{obj, "get", 1},"x", nil)
  assert(res == nil)

  local res = call(settable,{obj, "func", 1},"x", nil)
  assert(res == nil)


  nt()

  local res = call(settable,{obj, "get", print},"x")
  assert(res == 1)

  local res = call(settable,{obj, "func", print},"x")
  assert(res == 1)


end





-- Que testes vao ser realizados

if 1 then



if 1 then

test_simple()

test_wrong_parameters()



test_propget()

test_propput()



test_index_fb()

test_method_call_without_self()

end



if nil then

test_in_params()

test_out_params()

test_inout_params()

end



if 1 then

test_luadisp_in_params()

test_luadisp_out_params()

test_luadisp_inout_params()

test_luadisp_propget()

test_luadisp_propput()

end



if 1 then

test_iface_implementation()

test_connection_points()

end



if 1 then

test_safearrays()

end



if 1 then

test_stress()

end



if 1 then

test_NewObject()

test_DataTypes()

test_USERDEF_PTR()

end

test_field_redefinition()



end






exit(0)
















  print("\n=======> test_safearrays")



  teste = {}



  nt(1)


  teste.Teste = 10

  teste.TestSafeArrayIDispatch = function(self, array)

  end



  obj = luacom.ImplInterfaceFromTypelib(

    teste,"c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )



  assert(obj)



  array = {}



  local i = 1



  while i < 10 do

    array[i] = luacom.ImplInterfaceFromTypelib(

      teste, "c:\\users\\almendra\\testectl\\Debug\\teste.tlb", "ITeste" )



    i = i + 1

  end



  obj:TestSafeArrayIDispatch(array)

  obj = nil
  array = nil

  collectgarbage()




