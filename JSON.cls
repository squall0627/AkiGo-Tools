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
' �v���O�����h�c  : TECAEZ1
' ���V�X�e��   : �i�e�d�e�N�m���T�[�`�E���ƌn�Ɩ��x���V�X�e��
' �@�\��    : �G�N�Z���A�b�v���[�h�@�\
'
' �����T�v   : �G�N�Z�����[����AFTP�T�[�o�փA�b�v���[�h���A�T�[�o���Ńt�@�C���T�[�o�ֈړ�����B
'              �����l�́A�G�N�Z�����[�_�E�����[�h���ɁA�ėp�}�X�^�����ݒ肷��B
'
' ���쌠    : Copyrights (c) JFE Systems, lnc. All right reserved.
' �쐬��    : Chen Hao
' �쐬��    : 2016/02/12
'
' �C������
' 2016/02/12 Chen Hao : �V�K�쐬
''' *****************************************************************************
Option Explicit

Public result As Object

Private sc As Object
Private current_id As Long
Private max_id As Long

'�R���X�g���N�^
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

    '�R���X�g���N�^�ŁAJScript�I�u�W�F�N�g�𐶐�
    Set sc = CreateObject("ScriptControl")
    With sc
        .Language = "JScript"
        '�X�N���v�g���Ŏg�p����ϐ���錾
        .AddCode "var ary = new Array();"
        .AddCode "var row;"
        
        '�w�肵���C���f�b�N�X�A���̂̃f�[�^���擾����
        .AddCode "function GetValue(index, name) { return ary[index][name];}"
        
        '�z�񐔎擾�p
        .AddCode "function GetLength() { return ary.length;}"
        
        '�z���͗p
        .AddCode "function ParseArray(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var i=0;i < ja.length; i++ ) { " & _
                "dict.add(i, ParseObject(ja[i]));" & _
                "}" & _
                "return dict;}"
         
        '�I�u�W�F�N�g��͗p
        .AddCode "function ParseObject(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var key in ja) { " & _
                "if(ja[key] instanceof Array) { dict.add(key, ParseArray(ja[key]));}" & _
                "else if(ja[key] instanceof Object) { dict.add(key, ParseObject(ja[key]));}" & _
                "else {dict.add(key, ja[key]);}" & _
                "}" & _
                "return dict;}"
            
        '��̓��C�����\�b�h
        .AddCode "function ParseJson() {" & _
                "if(ary instanceof Array) { return ParseArray(ary);}" & _
                "else { return ParseObject(ary);}" & _
                "}"

        '�f�[�^�ǉ��p
        '.AddCode "function NewRow() { row = {}; ary.append(row); }"
        .AddCode "function NewRow() { row = new Object(); ary[ary.length] = row; }"
        .AddCode "function AddData(key, value) { row[key] = value; }"
        .AddCode "function EncodeData() { return '[' + EncodeDataDetails(ary, '') + ']'; }"
        .AddCode code
    End With
    
    current_id = -1
    max_id = 0
End Sub

'JSON�`���̃f�[�^����͂���
Public Sub Parse(ByRef data As String)
    'ary�Ƃ����I�u�W�F�N�g�Ɏ擾����JSON�`���̃f�[�^��W�J
    sc.AddCode "var ary = " & data & ";"
    
    '�z�񐔂��m��
    max_id = sc.CodeObject.GetLength("")
    
    'JSON��͂���Dictionary�ɕێ�
    Set result = sc.CodeObject.ParseJson("")
    
End Sub

'�܂��f�[�^�����݂��邩
Public Function HasNext() As Boolean
    current_id = current_id + 1
    HasNext = (current_id < max_id)
End Function

'�C���f�b�N�X���w�肵�Ẵf�[�^�擾
Public Function GetValueAt(ByVal index As Long, ByVal id As String) As String
    GetValueAt = sc.CodeObject.GetValue(index, id)
End Function

'�J�����g�s�̃f�[�^�擾
Public Function GetValue(ByVal id As String) As String
    GetValue = GetValueAt(current_id, id)
End Function


'�V�����s��錾
Public Function NewRow()
    Call sc.CodeObject.NewRow("")
End Function

'�f�[�^���s�ɒǉ�
Public Function AddData(ByVal key As String, ByVal value As String)
    Dim escapeStr As String
    '\��\\"�ɕϊ�����Add
    escapeStr = Replace(value, "\", "\\")
    '"��\"�ɕϊ�����Add
    escapeStr = Replace(escapeStr, Chr(34), "\" & Chr(34))
    Call sc.CodeObject.AddData(key, escapeStr)
    'sc.Eval ("row.key = value;")
End Function

'JSON�`���ɕϊ�
Public Function encode() As String
    encode = sc.CodeObject.EncodeData("")
End Function

'�f�X�g���N�^
Public Sub Class_Terminate()
    Set sc = Nothing
End Sub