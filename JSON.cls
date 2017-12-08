VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' *****************************************************************************
' プログラムＩＤ  : TECAEZ1
' 情報システム   : ＪＦＥテクノリサーチ・現業系業務支援システム
' 機能名    : エクセルアップロード機能
'
' 処理概要   : エクセル帳票から、FTPサーバへアップロードし、サーバ側でファイルサーバへ移動する。
'              埋込値は、エクセル帳票ダウンロード時に、汎用マスタ等より設定する。
'
' 著作権    : Copyrights (c) JFE Systems, lnc. All right reserved.
' 作成者    : Chen Hao
' 作成日    : 2016/02/12
'
' 修正履歴
' 2016/02/12 Chen Hao : 新規作成
''' *****************************************************************************
Option Explicit

Public result As Object

Private sc As Object
Private current_id As Long
Private max_id As Long

'コンストラクタ
Public Sub Class_Initialize()

    Dim code As String
    
    code = ""
    code = code & "function EncodeDataDetails(obj,str) {"
    code = code & " var result='';"
    code = code & " for (key in obj){"
    code = code & "     if (typeof obj[key] =='object'){"
    code = code & "         result += '{'+EncodeDataDetails( obj[key],str ) +'},';"
    code = code & "     } else {"
    code = code & "         result += '""' + key + '"": ""'+ obj[key] +'"",';"
    code = code & "     }"
    code = code & " }"
    code = code & " return result.slice(0,result.length-1);"
    code = code & "}"

    'コンストラクタで、JScriptオブジェクトを生成
    Set sc = CreateObject("ScriptControl")
    With sc
        .Language = "JScript"
        'スクリプト内で使用する変数を宣言
        .AddCode "var ary = new Array();"
        .AddCode "var row;"
        
        '指定したインデックス、名称のデータを取得する
        .AddCode "function GetValue(index, name) { return ary[index][name];}"
        
        '配列数取得用
        .AddCode "function GetLength() { return ary.length;}"
        
        '配列解析用
        .AddCode "function ParseArray(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var i=0;i < ja.length; i++ ) { " & _
                "dict.add(i, ParseObject(ja[i]));" & _
                "}" & _
                "return dict;}"
         
        'オブジェクト解析用
        .AddCode "function ParseObject(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var key in ja) { " & _
                "if(ja[key] instanceof Array) { dict.add(key, ParseArray(ja[key]));}" & _
                "else if(ja[key] instanceof Object) { dict.add(key, ParseObject(ja[key]));}" & _
                "else {dict.add(key, ja[key]);}" & _
                "}" & _
                "return dict;}"
            
        '解析メインメソッド
        .AddCode "function ParseJson() {" & _
                "if(ary instanceof Array) { return ParseArray(ary);}" & _
                "else { return ParseObject(ary);}" & _
                "}"

        'データ追加用
        '.AddCode "function NewRow() { row = {}; ary.append(row); }"
        .AddCode "function NewRow() { row = new Object(); ary[ary.length] = row; }"
        .AddCode "function AddData(key, value) { row[key] = value; }"
        .AddCode "function EncodeData() { return '[' + EncodeDataDetails(ary, '') + ']'; }"
        .AddCode code
    End With
    
    current_id = -1
    max_id = 0
End Sub

'JSON形式のデータを解析する
Public Sub Parse(ByRef data As String)
    'aryというオブジェクトに取得したJSON形式のデータを展開
    sc.AddCode "var ary = " & data & ";"
    
    '配列数を確定
    max_id = sc.CodeObject.GetLength("")
    
    'JSON解析してDictionaryに保持
    Set result = sc.CodeObject.ParseJson("")
    
End Sub

'まだデータが存在するか
Public Function HasNext() As Boolean
    current_id = current_id + 1
    HasNext = (current_id < max_id)
End Function

'インデックスを指定してのデータ取得
Public Function GetValueAt(ByVal index As Long, ByVal id As String) As String
    GetValueAt = sc.CodeObject.GetValue(index, id)
End Function

'カレント行のデータ取得
Public Function GetValue(ByVal id As String) As String
    GetValue = GetValueAt(current_id, id)
End Function


'新しい行を宣言
Public Function NewRow()
    Call sc.CodeObject.NewRow("")
End Function

'データを行に追加
Public Function AddData(ByVal key As String, ByVal value As String)
    Dim escapeStr As String
    '\を\\"に変換してAdd
    escapeStr = Replace(value, "\", "\\")
    '"を\"に変換してAdd
    escapeStr = Replace(escapeStr, Chr(34), "\" & Chr(34))
    Call sc.CodeObject.AddData(key, escapeStr)
    'sc.Eval ("row.key = value;")
End Function

'JSON形式に変換
Public Function encode() As String
    encode = sc.CodeObject.EncodeData("")
End Function

'デストラクタ
Public Sub Class_Terminate()
    Set sc = Nothing
End Sub
