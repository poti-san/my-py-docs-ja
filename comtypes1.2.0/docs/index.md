---
parent: comtypes 1.2.0
title: comtypesパッケージ
---
{% raw %}
# comtypesパッケージ

**comtypes**はctypesライブラリーを基礎とした*純Python*なCOMパッケージです。**ctypes**はPython 2.5以降に同梱されており、Python 2.4でも追加でダウンロードできます。

**pywin32**パッケージは*ディスパッチベース*なCOMインターフェイスの優れたクライアントサイド機能をサポートしますが、C++コードでラップしなければ*カスタム*COMインターフェイスにアクセスできません。

**comtypes**パッケージはカスタムとディスパッチベース両方のCOMインターフェイスへのアクセスと実装を簡単にします。

このドキュメントは**comtypes** 1.1.11について記述しています。

新機能：comtypesを用いたCOMサーバー実装文書の入り口は[comtypes_server](server.html)です。

## comtypes.clientパッケージ

**comtypes.client**パッケージは**comtypes**の高水準機能を実装します。

### COMオブジェクトの作成とアクセス

**comtypes.client**はCOMオブジェクトの作成とアクセスを可能とする３つの関数を公開します。

`CreateObject(progid, clsctx=None, machine=None, interface=None, dynamic=False, pServerInfo=None)`

COMオブジェクトを作成してインターフェイスのポインターを返します。

`progid`は作成するオブジェクトを指定します。`"InternetExplorer.Application"`や`"{2F7860A2-1473-4D75-827D-6C4E27600CAC}"`のような文字列、`comtypes.GUID`インスタンス、`comtypes.GUID`インスタンスかGUID文字列を`_clsid_`属性に持つ任意のオブジェクトを指定できます。

`clsctx`はオブジェクトの作成方法を指定します。`comtypes.CLSCTX_...`定数の任意の組み合わせで、未指定時は`comtypes.CLSCTX_SERVER`です。

`machine`はオブジェクトを作成する別のコンピューターを指定します。コンピューターの名前かIPアドレスを表す文字列です。DCOMはこの機能で有効化します。

`inyerface`は返されるインターフェイスの種類を指定します。未指定の場合、**comtypes**が便利なインターフェイスを決定します。

`dynamic`は作成するインターフェイスが動的ディスパッチを使用すべきか指定します。オートメーションのインターフェイスでのみ有効であり、タイプライブラリーのラッパーは作成しません。

`pServerInfo`は`machine`引数より詳細なリモートコンピューターの情報を指定します。`COSERVERINFO`のポインターです。`machine`と`pServerInfo`は同時に指定できません。DCOMはこの機能で有効化します。

`CoGetObject(displayname, interface=None)` 

名前付きCOMオブジェクトを作成してインターフェイスのポインターを返します。`displayname`の解釈にはMicrosoftドキュメントの`CoGetObject`関数を参照してください。例えば`"winmgmts:"`は[WMIモニカー](http://www.microsoft.com/technet/scriptcenter/guide/sas_wmi_jgfx.mspx?mfr=true)用の`displayname`です。

```python
wmi = CoGetObject("winmgmts:")
```

`interface`と`dynamic`は`CreateObject`関数と同じ意味です。

`GetActiveObject(progid, interface=None)`  

起動中のオブジェクトのポインターを返します。`progid`はOLE登録データベース中のアクティブオブジェクトを指定します。

関数はCOMオブジェクトが起動中でCOM ROT（COM実行中オブジェクトテーブル）に登録済みの場合に成功します。これはすべてのCOMオブジェクトには当てはまりません。この引数の説明は`CreateObject`関数を参照してください。

以上の３つの関数はオブジェクトが型情報を提供する場合にタイプライブラリーのラッパーを自動的に作成します。オブジェクトがタイプライブラリーを公開しない場合、`GetModule`関数を呼び出してラッパーを作成できます。

### COMオブジェクトの使用

作成関数（`CreateObject`、`CoGetObject`、`GetActiveObject`）の返すCOMインターフェイスポインターはインターフェイスのメソッドとプロパティを公開します（`dynamic`が渡されていない限り）。

`comtypes`はCOMインターフェイスの事前バインド（early binding）を用いるため、インターフェイスのメソッドとプロパティは調査可能です。Pythonの組み込み`help`関数で概要を取得できます。

`MSScriptControl.ScriptControl`はMicrosoftスクリプトエンジンのProgIDです。これは興味深いCOMオブジェクトであり、JScriptやVBScriptのプログラムを実行できます[こちら](scriptcontrol.html)で以下のコマンドの完全な出力を確認できます。

```py
>>> from comtypes.client import CreateObject
>>> engine = CreateObject("MSScriptControl.ScriptControl")
>>> help(engine)
.....
>>>
```

#### メソッドの呼び出し

COMメソッドは他のPythonオブジェクトと同様に呼び出せます。位置または名前付き引数も使用できます。

IDLで`[out]`や`[out, retval]`を指定された引数はメソッドの成功時にタプルとして返されます。それらがない場合、`HRESULT`値が返されます。メソッドが成功して`[out]`や`[out, retval]`引数の値が返されたとき、`HRESULT`値は消失します。

COMメソッドの呼び出しに失敗した場合、`HRESULT`値を持つ`COMError`例外が発生します。

#### プロパティのアクセス

COMプロパティにはいくつかの難題があります。プロパティには読み書き、読み取り専用、書き込み専用があります。引数は0個、1個、複数個で、省略可能な場合もあります。

引数の無いプロパティは通常の方法でアクセスできます。以下の例はInternet Explorerの`Visible`プロパティのデモです。

```py
>>> ie = CreateObject("InternetExplorer.Application")
>>> print ie.Visible
False
>>> ie.Visible = True
>>>
```

##### 引数のあるプロパティ（名前付きプロパティ）

引数のあるプロパティは添字表記でアクセスできます。以下の例はExcelを開始して、新しいワークブックを作り、`xlRangeValueDefault`書式のあるセルの内容にアクセスします（Office 2003で動作確認しています）。

```py
>>> xl = CreateObject("Excel.Application")
>>> xl.Workbooks.Add()
>>> from comtypes.gen.Excel import xlRangeValueDefault
>>> xl.Range["A1", "C1"].Value[xlRangeValueDefault] = (10,"20",31.4)
>>> print xl.Range["A1", "C1"].Value[xlRangeValueDefault]
(10, "20", 31.4)
>>>
```

##### 省略可能引数のあるプロパティ

Excelタイプライブラリー（または*comtypes.gen*で作成されたラッパーモジュール）を探ると`.Value`プロパティの引数が省略可能であることに気づくでしょう。`xlRangeValueDefault`引数を渡さなくても（あるいは知ってさえいなくても）取得／設定できます。

不幸にもPythonでは引数のない添字表記は使えません。

```py
>>> xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
  File "<stdin>", line 1
    xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
                               ^
SyntaxError: invalid syntax
>>> print xl.Range["A1", "C1"].Value[]
  File "<stdin>", line 1
    print xl.Range["A1", "C1"].Value[]
                                     ^
SyntaxError: invalid syntax
>>>
```

従って、**comtypes**はこれらのプロパティにアクセスするいくつかの方法を提供すべきです。引数を渡さずに名前付きプロパティを取得するには、プロパティの呼び出しが使えます。

```py
>>> print xl.Range["A1", "C1"].Value()
(10, "20", 31.4)
>>>
```

空のスライスまたはタプルを添字に使っても良いです。

```py
>>> print xl.Range["A1", "C1"].Value[:]
(10, "20", 31.4)
>>> print xl.Range["A1", "C1"].Value[()]
(10, "20", 31.4)
>>>
```

引数を渡さずに名前付きプロパティを設定するには、空のスライスまたはタプルが使えます。

```py
>>> xl.Range["A1", "C1"].Value[:] = (3, 2, 1)
>>> xl.Range["A1", "C1"].Value[()] = (1, 2, 3)
>>>
```

#### lcid引数

いくつかのCOMメソッドとプロパティは省略可能な`lcid`引数を持ちます。言語識別子の指定に使います。作成されたモジュールではこの引数は常に0を渡します。これが期待する動作ではない場合、作成されたコードを編集してください。

#### データ型の変換

**comtypes**は通常、期待されるだろう方法でCOMとPythonの引数と戻り値を変換します。

`VARIANT`引数は特殊な処理が必要な場合があります。`VARIANT`は多数の型を保持できます。整数、単精度浮動小数点、文字列のような単純な型ばかりでなく、より複雑な単次元・複数次元の配列もです。`VARIANT`の保持する値はcomtypesが自動的に割り当てる*typecode*により特定されます。

単純なシーケンス（リストやタプル）を`VARIANT`引数として渡したとき、COMサーバーは`typecode`に`VT_ARRAY | VT_VARIANT`を指定して`VARIANT`の`SAFEARRAY`を保持する`VARIANT`を受け取ります。

しかし、COMサーバーのメソッドにはこのような配列を許容せず、例えば`typecode`が`VT_ARRAY | VT_I2`の`short`整数配列、`VT_ARRAY | VT_INT`の整数配列、`VT_ARRAY | VT_BSTR`の文字列配列を必要とするものがあります。

これらの`VARIANT`を作成するには、COMメソッドに適したPython型コードを指定して作成したPythonの`array.array`インスタンスを渡します。NumPy配列を使うこともできます。詳細は下記のセクションで記述します。

`array.array`型コードから`VARIANT`の型コードへのマップは`comtypes.automation`モジュールで辞書として定義されています。

```python
_arraycode_to_vartype = {
    "b": VT_I1,
    "h": VT_I2,
    "i": VT_INT,
    "l": VT_I4,

    "B": VT_UI1,
    "H": VT_UI2,
    "I": VT_UINT,
    "L": VT_UI4,

    "f": VT_R4,
    "d": VT_R8,
}
```

例えばAutoCADは引数に型コード`VT_ARRAY | VT_I2`や`VT_ARRAY | VT_R8`の`VARIANT`を必要とするCOMサーバーのひとつです。次のコードはユーザーが提供してくれたものです。

```python
""" comtypeでAutoCADを自動化する方法のサンプルデモ：
点と線の描画への追加、それらへの異なる型のxdataのアタッチ。
目的はcomtypesを使って異なる型のVARIANTを作成する方法を実際に示すことです。
それらのVARIANTはAutoCAD COM APIの多くのメソッドで必要です。
次のコードのテストにはAutoCADの起動が必要です。 """

import array
import comtypes.client

# AutoCADアプリケーションの実行中インスタンスを取得します。
app = comtypes.client.GetActiveObject("AutoCAD.Application")

# ModelSpaceオブジェクトを取得します。
ms = app.ActiveDocument.ModelSpace

# ModelSpaceに点を追加します。
pt = array.array('d', [0,0,0])
point = ms.AddPoint(pt)

# ModelSpaceに線を追加します。
pt1 = array.array('d', [1.0,1.0,0])
pt2 = array.array('d', [2.0,2.0,0])
line = ms.AddLine(pt1, pt2)

# 点に整数型xdataを追加します。
point.SetXData(array.array("h", [1001, 1070]), ['Test_Application1', 600])

# 線に倍精度浮動小数点型xdataを追加します。
line.SetXData(array.array("h", [1001, 1040]), ['Test_Application2', 132.65])

# 線に文字列型xdataを追加します。
line.SetXData(array.array("h", [1001, 1000]), ['Test_Application3', 'TestData'])

# 線にリスト型（この場合は点の座標群）を追加します。
line.SetXData(array.array("h", [1001, 1010]),
          ['Test_Application4', array.array('d', [2.0,0,0])])

print "Done."
```

### NumPy相互運用

NumPyはPythonの配列における事実上の標準です。comtypesの使用にNumPyは必須ではありませんが、comtypesはNumPyとの相互運用に多彩なオプションを提供します。この機能を完全に使うにはNumPyバージョン1.7以上が必要です。

#### 入力引数としてのNumPy配列

NumPy配列は`VARIANT`配列引数に渡せます。この配列は型に応じた`SAFEARRAY`に変換されます。型変換は`numpy.ctypeslib`モジュールで定義されています。次の表はNumPy配列から`SAFEARRAY`への（ほぼ）直接変換で高速に動作する型変換を示しています。この表にない型の配列でも項目毎の変換は可能です。

| NumPy型                              | VARIANT型 |
| ----------------------------------- | -------- |
| `int8`                              | VT_I1    |
| `int16`, `short`                    | VT_I2    |
| `int32`, `int`, `intc`, `int_`      | VT_I4    |
| `int64`, `long`, `longlong`, `intp` | VT_I8    |
| `uint8`, `ubyte`                    | VT_UI1   |
| `uint16`, `ushort`                  | VT_UI2   |
| `uint32`, `uint`, `uintc`           | VT_UI4   |
| `uint64`, `ulonglong`, `uintp`      | VT_UI8   |
| `float32`                           | VT_R4    |
| `float64`, `float_`                 | VT_R8    |
| `datetime64`                        | VT_DATE  |

#### 出力引数としてのNumPy配列

既定では、comtypesは`SAFEARRAY`出力引数を項目毎に変換してPythonオブジェクトのタプルを作成します。大きな`SAFEARRAY`を扱う場合、この変換はコストがかかります。comtypesはこの動作をNumPy配列の返却に変更する`comtypes.safearray`で`safearray_as_ndarray`コンテキストマネージャーを提供しています。この代替動作は`SAFEARRAY`のメモリを`ndarray`にコピーするため、項目毎にPythonを呼び出すよりも高速です。失敗時、NumPy配列は項目毎のコピーで作成されます。`safearray_as_ndarray`コンテキストマネージャーはスレッドセーフです。あるスレッドでの使用は別スレッドの動作に影響しません。

次のコードは`safearray_as_ndarray`コンテキストマネージャーを使用した架空の例です。タプルの代わりにNumPy配列を取得するために任意のプロパティやメソッド呼び出しで使用できます。

```python
""" safearray_as_ndarrayコンテキストマネージャーの使い方のサンプルデモ """

from comtypes.safearray import safearray_as_ndarray

# 戻り値のSAFEARRAYをタプルとして返す架空の例。
data1 = some_interface.some_property

# これはNumPy配列を返します。基本型では上記より高速でしょう。
with safearray_as_ndarray:
    data2 = some_interface.some_property
```

### COMイベント;

いくつかのCOMオブジェクトはイベントをサポートします。イベントは何かが起きたときにユーザーへ知らせることを可能にします。COMの標準機能は「コネクションポイント」と呼ばれる機能に基づきます。

注意：イベントハンドラーの実装時に読むべきルールはcomtypesサーバードキュメントの[COMメソッドの実装（未翻訳）](server.html#implementing-com-methods)セクションに記載があります。

`GetEvents(source, sink, interface=None)`  

イベントシンクとCOMオブジェクト`source`を接続します。

イベントは`sink`オブジェクトのメソッドを呼び出します。メソッドの名前は`interfacename_methodname`か`mmethodname`の必要があります。メソッドは`this`引数、そしてイベントの持つ任意の追加引数を伴って呼び出されます。

`interface`は`source`オブジェクトの外向きインターフェイスです。**comtypes**が`source`の外向きインターフェイスを決定できない場合に指定すべきです。

`GetEvents`はアドバイザリ接続を返します。イベントを受信したい間、接続を維持する必要があります。アドバイザリ接続を破棄するには単純に削除します。

`ShowEvents(source, interface=None)`
イベントシンクを構築してデバッグ用に`source`オブジェクトと接続します。最初にイベントシンクは外向きインターフェイスで見つかったすべてのイベント名を出力します。以降はイベントの発生次第、そのイベントと引数を出力します。`ShowEvents`の返す接続オブジェクトはイベントを受信したい限り生存を維持してください。オブジェクトを削除すると`source`オブジェクトへの接続は削除されます。

実際にイベントを受信するには`PumpEvents`を呼び出してCOMを正確に機能させます。

`PumpEvents(timeout)`  

COMの正常動作に必要な方法である時間待機します。シングルスレッド・アパートメントではWindowsのメッセージループを回し、マルチスレッド・アパートメントでは単純に待機します。`timeout`引数では浮動小数点数により秒未満の時間を指定できます。

Control+Cを押すと`KeyboardError`例外が発生して関数は即座に終了します。

#### 具体例

ここではExcelからイベントを検索・受信する方法をデモします。

```py
>>> from comtypes.client import CreateObject
>>> xl = CreateObject("Excel.Application")
>>> xl.Visible = True
>>> print xl
<POINTER(_Application) ptr=0x29073c at c156c0>
>>>
```

`ShowEvents`関数はインタラクティブPythonインタプリターでオブジェクトのイベントに取り掛かる便利なヘルパーです。

`ShowEvents`を呼び出してExcelの送信するイベントに接続できます。`ShowEvents`は最初に`_Application`オブジェクトに存在するイベントを列挙します。

```py
>>> from comtypes.client import ShowEvents
>>> connection = ShowEvents(xl)
# event found: AppEvents_WorkbookSync
# event found: AppEvents_WindowResize
# event found: AppEvents_WindowActivate
# event found: AppEvents_WindowDeactivate
# event found: AppEvents_SheetSelectionChange
# event found: AppEvents_SheetBeforeDoubleClick
# event found: AppEvents_SheetBeforeRightClick
# event found: AppEvents_SheetActivate
# event found: AppEvents_SheetDeactivate
# event found: AppEvents_SheetCalculate
# event found: AppEvents_SheetChange
# event found: AppEvents_NewWorkbook
# event found: AppEvents_WorkbookOpen
# event found: AppEvents_WorkbookActivate
# event found: AppEvents_WorkbookDeactivate
# event found: AppEvents_WorkbookBeforeClose
# event found: AppEvents_WorkbookBeforeSave
# event found: AppEvents_WorkbookBeforePrint
# event found: AppEvents_WorkbookNewSheet
# event found: AppEvents_WorkbookAddinInstall
# event found: AppEvents_WorkbookAddinUninstall
# event found: AppEvents_SheetFollowHyperlink
# event found: AppEvents_SheetPivotTableUpdate
# event found: AppEvents_WorkbookPivotTableCloseConnection
# event found: AppEvents_WorkbookPivotTableOpenConnection
# event found: AppEvents_WorkbookBeforeXmlImport
# event found: AppEvents_WorkbookAfterXmlImport
# event found: AppEvents_WorkbookBeforeXmlExport
# event found: AppEvents_WorkbookAfterXmlExport
>>> print connection
<comtypes.client._events._AdviseConnection object at 0x00C16AD0>
>>>
```

`ShowEvents`の戻り値を`connection`変数に割り当てています。これによりExcelへの接続は生存を維持され、発生したイベントを出力できます。

COMイベントの正確な受信にはメッセージループの実行が重要です。`PumpEvents()`関数は一定時間だけそれを実行します。以下はこの関数の呼び出して、その間にExcelワークシートをインタラクティブに開いたときに起きた内容です。`comtypes`はイベントを実行時の引数と共に出力します。

```py
>>> from comtypes.client import PumpEvents
>>> PumpEvents(30)
Event AppEvents_WorkbookOpen(None, <POINTER(_Workbook) ptr=...>)
Event AppEvents_WorkbookActivate(None, <POINTER(_Workbook) ptr=...>)
Event AppEvents_WindowActivate(None, <POINTER(Window) ptr=...>, <POINTER(_Workbook) ptr=...>)
>>>
```

最初の引数は常に`this`ポインターで、`comtypes`の内部事情により常に`None`が渡されます。他の引数はイベントに依存します。接続の削除には単に`connection`変数を削除します。これはPythonのガベージコレクタを呼び出して接続を即座に削除するために必要です。 削除するとExcelのイベントは受信されません。

```py
>>> del connection
>>> import gc; gc.collect()
123
>>>
```

自作のコードでイベントを処理したい場合、上とよく似た方法で`GetEvents()`関数を使えます。この関数は最初の引数にCOMオブジェクト、二つ目の引数にイベントを処理するイベントシンクPythonオブジェクトを受け取ります。イベントシンクは処理するイベントと似た名前のメソッドを持つべきです。処理したいイベントに対応するメソッドだけ実装します。他のイベントは無視されます。

次のコードは`AppEvents_WorkbookOpen`イベントを処理するクラスを定義して、そのクラスのインスタンスを作成して、`GetEvents()`関数の2番目の引数に渡しています。

```py
>>> from comtypes.client import GetEvents
>>> class EventSink(object):
...     def AppEvents_WorkbookOpen(self, this, workbook):
...         print "WorkbookOpened", workbook
...         # add your code here
...
>>> sink = EventSink()
>>> connection = GetEvents(xl, sink)
>>> PumpEvents(30)
WorkbookOpened <POINTER(_Workbook) ptr=0x291944 at 1853120>
>>>
```

イベントハンドラのメソッドはCOMメソッドの実装と同じcomtypesの呼び出し規約に対応させることに注意してください。詳細は[COMメソッドの実装](server.html#implementing-com-methods)を読んでください。

### タイプライブラリー

#### タイプライブラリーへのアクセス

**comtypes**はカスタムCOMインターフェイスでも事前バインディングを使います。`comtypes.IUnknown`クラスを継承するPythonクラスを記述してください。このクラスはIDL表記と似た方法でインターフェイスのメソッドとプロパティを記述します。

インターフェイスクラスは手動でも記述できますが、幸運にも**comtypes**はコードジェネレータを含みます。コードジェネレータはCOMタイプライブラリーから自動的にPythonのインターフェイスクラス等を含むモジュールを作成します。

`GetModule(tlib)`

> COMタイプライブラリーのPythonラッパーを作成します。COMオブジェクトがタイプ情報を公開する場合、関数はそのオブジェクトの作成時に自動で呼び出されます。
> 
> `tlib`は読み込み済みのタイプライブラリーの`ITypeLib`COMポインター、タイプライブラリーを含むファイルのパス（.tlb;.exe;.dll）、タイプライブラリーのGUID・メジャーおよびマイナーバージョン番号・オプションでLCIDを含むタプルまたはリスト、タイプライブラリーを特定する\_\_reg_libid\_\_と\_\_reg_version\_\_属性を持つ任意のオブジェクトです。
> 
> `GetModule(tlib)`はタイプライブラリーからインターフェイス、coclass、定数、構造体、モジュールオブジェクトを含むPythonモジュールを作成して、モジュールオブジェクト自体を返します。モジュールは`comtypes.gen`パッケージ内部に作成されます。モジュール名はタイプライブラリーのGUID、バージョン数、LCIDから作成されます。Pythonモジュール名として有効であり、`import`ステートメントでインポートできます。2つめのラッパモジュールも`comtypes.gen`パッケージにタイプライブラリーの「名前」からより短い名前で作成されます。このモジュールは実際のラッパーモジュールからすべてをインポートしますが、モジュール名が打ちやすいのでより簡単にインポートできます。
> 
> 具体的にInternet Explorerのタイプライブラリーは`SHDocVw`という名前（IDLファイルで特定された名前であり、ファイル名ではありません）、`{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}`というGUID、`1.1`というバージョン数を持ちます。これに対する実際のタイプライブラリーラッパーモジュールの名前は`comtypes.gen._EAB22AC0_30C1_11CF_A7EB_0000C05BAE0B_0_1_1`であり、2つめのラッパーモジュールの名前は`comtypes.gen.SHDocVw`です。
> 
> スクリプトをpy2exeで凍結するとき、次の記述でp2exeがそれらのタイプライブラリーラッパーを含むことを保証できます。
> 
> ```python
> import comtypes.gen.SHDocVw
> ```
> 
> どこかで。

`gen_dir`

> この変数はタイプライブラリーラッパーが書き込まれるディレクトリを決定します。`None`の場合、モジュールはメモリ上に作成されます。
> 
> `comtypes.client.gen_dir`は`comtypes.client`モジュールの初回インポート時に計算されます。これが有効なファイルシステムパスである場合、`comtypes.gen`パッケージのディレクトリに設定されます。それ以外の場合は`None`に設定されます。
> 
> py2exeで凍結されたスクリプトでは、**comtypes.gen**のディレクトリはZIPアーカイブのどこかで、`gen_dir`は`None`です。タイプライブラリーラッパーは実行時に作成されてもファイルシステムへは書き込まれません。モジュールはメモリ上でのみ作成されます。
> 
> `comtypes.client.gen_dir`はタイプライブラリーラッパーのファイルシステムへの書き込みを抑制するために`None`に設定できます。それより下側では大きなタイプライブラリーのコード生成にはいくらかの時間が必要となります。

#### 具体例

`GetModule`関数を使ってInternet Explorerのタイプライブラリーラッパーモジュールを作成するいくつかの方法を紹介します。

```py
>>> from comtypes.client import GetModule
>>> GetModule("shdocvw.dll")
>>> GetModule(["{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}", 1, 1)
>>>
```

次のコード断片は実行するとInternet Explorerタイプライブラリーラッパーモジュールを自動作成します。スクリプトがpy2exeで凍結されるときはそのモジュールを実行ファイルに含めます。

```py
>>> import sys
>>> if not hasattr(sys, "frozen"):
>>>     from comtypes.client import GetModule
>>>     GetModule("shdocvw.dll")
>>> import comtypes.gen.ShDocVw
>>>
```

#### （アルファベットの）大文字と小文字の区別

基本的にCOMは大文字と小文字を区別しない技術です（おそらくVisual Basicのためです）。しかし、IDLファイルから作成されたタイプライブラリーは識別子の大文字小文字を維持しないことがあります。具体例は[http://support.microsoft.com/kb/220137](http://support.microsoft.com/kb/220137)を参照してください。

Python（とC/C++）は大文字小文字を区別する言語であり、**comtypes**も同様です。従って、`obj.QueryInterface(...)`と書くべきであり、`obj.queryinterface(...)`とは書けません。

タイプライブラリー（用に作成されたPythonモジュール）の識別子の大文字と小文字がIDLファイルと異なる場合に生じる問題に対処するため、**comtypes**はCOMインターフェイスのメソッドとプロパティへのアクセスで大文字と小文字を区別しない属性を許します。この動作はPythonのCOMインターフェイスで`_case_insensitive_`属性を`True`に設定すると有効になります。COMインターフェイス由来の場合、大文字と小文字の区別はインターフェイス毎に有効化・無効化できます。

`GetModule`関数で作成したコードではこの属性は`True`に設定されます。大文字と小文字を区別しないアクセスはパフォーマンスに少しのペナルティがあります。これを避けたい場合、作成されたコードを編集して`_case_insenstive_`属性を`False`に設定します。

### スレッド処理

誰かがシングルスレッドアパートメント、マルチスレッドアパートメント、`sys.coinit_flags`、`CoInitialize`、`CoUninitialize`等に言及します。これらはずっと将来に記述されます。

誰かはスレッド処理問題、メッセージループにも言及します。

### その他の事項

誰かが`logging`、`gen_dir`、`warp`、`_manage`（？）を記述します。

### リンク

Yaroslav Kourovtsevによる[「Working with custom COM interfaces from Python」](http://www.codeproject.com/KB/COM/python-comtypes-interop.aspx)は**comtypes**でカスタムCOMオブジェクトにアクセスする方法を記載しています。

## ダウンロード

**comtypes**プロジェクトは[GitHub](https://github.com/enthought/comtypes)でホストされています。公開版は[GitHubのReleasesセクション](https://github.com/enthought/comtypes/releases)からダウンロードできます。

{% endraw %}
