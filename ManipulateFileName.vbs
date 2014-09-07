Option Explicit
'================================================================
'�t�@�C��������X�N���v�g ManipulateFileName.vbs
'================================================================
'	�w��t�H���_�z���̃t�@�C�����ׂĂɑ΂��āA�t�@�C�����̕�������ꊇ�ŕύX���܂��B
'	���[�h�w��ɂ��A�w�蕶������t�@�C�����̐擪�������͖����ɕt��������
'	�t�@�C��������w�蕶������������邱�Ƃ��o���܂��B
'
'�y�����z
'	�ϊ��Ώۃt�@�C���̂���t�H���_�p�X
'	�t��������or����������������
'	���[�h
'		1:�擪�ɕ�����t��
'		2:�����ɕ�����t��
'		3:�w�蕶������폜)
'
'�y�߂�l�z
'   OK  0	����I��
'   NG  1	�ُ�I��

'�y���O�o�́z
'	�������s�A�I�����ƁA�G���[����ьx�������������ꍇ�́A���s�t�H���_�Ƀ��O���o�͂��܂��B
'
'	�y�t�H�[�}�b�g�FCSV�`���z
'	"��������(YYYY/MM/DD hh:mm:ss)","�X�e�[�^�X","���b�Z�[�W"
'
'	�y�X�e�[�^�X�̎�ށz
'		Info	(���)
'		Warning	(�x��)
'		Error	(�G���[)
'
'
'----------------------------------------------------------------
'�@�萔
'----------------------------------------------------------------

	'�������b�Z�[�W
	Const MSG_START		=	"�������J�n���܂��B"
	Const MSG_END		=	"�������I�����܂����B"

	'�G���[���b�Z�[�W
	Const ERR_MSG_ARG						=	"�s���ȃp�����[�^���w�肳��܂����B"
	Const ERR_MSG_NOT_FOLDER_PATH			=	"�w�肳�ꂽ�p�X�̓t�H���_�ł͂���܂���B"
	Const ERR_MSG_FILE_NOT_OPENN			=	"�t�@�C���I�[�v���Ɏ��s���܂����B"
	Const ERR_MSG_CANNOT_FILE_NAME			=	"�t�@�C�����Ɏg�p�ł��Ȃ��������w�肳��܂����B"

	'�x�����b�Z�[�W
	Const WARNING_MSG_FILE_SAME_NAME			=	"����t�@�C�����̃t�@�C�������ɑ��݂��܂��B"

	'�X�e�[�^�X�̎��
	Const STATUS_TYPE_INFO		=	"Info"
	Const STATUS_TYPE_WARNING	=	"Warning"
	Const STATUS_TYPE_ERROR		=	"Error"

	'�L��
	Const STR_YEN					=	"\"
	Const STR_DOUBLEQUOT			=	""""
	Const STR_LEFT_SQUARE_BRACKET	=	"["
	Const STR_RIGHT_SQUARE_BRACKET	=	"]"
	Const STR_JOIN_LOG				= 	""","""
	
	'�f�t�H���g���O�o�̓p�X
	Const STR_LOG_FILE_NAME		=	"ManipulateFileName.log"
	
	'�����񑀍샂�[�h
	Const STR_MODE_FROMT	=	"1"		'�擪�ɕt��
	Const STR_MODE_REAR		=	"2"		'�����ɕt��
	Const STR_MODE_DELETE	=	"3"		'�폜
	
'----------------------------------------------------------------
'�@�X�e�[�^�X�萔�錾
'----------------------------------------------------------------
	Const OK	= "0"	'����
	Const NG	= "1"	'�ُ�
'----------------------------------------------------------------

Dim objFS
Dim STR_MANIPULATE_MODE		'�����񑀍샂�[�h

'--- FileSystemObject�̐��� ---
Set objFS = CreateObject("Scripting.FileSystemObject")
STR_MANIPULATE_MODE	=	0


Call Main()



'----------------------------------------------------------------
'���C���������s���܂��B
'
'@param		����
'@return	RETURN_TRUE		����
'			RETURN_FALSE	�G���[or�x��
'----------------------------------------------------------------
Function Main()
	Dim objArgs			'�����擾�p�I�u�W�F�N�g
	Dim strErrMsg		'�G���[���b�Z�[�W
	Dim intMode			'�����񑀍샂�[�h
	Dim strTargetDir	'�Ώۃf�B���N�g��
	Dim strManipulate	'���삷�镶����
	Dim strLogPath		'���O�t�@�C���p�X
	Dim objParam

	'--- ������ ---
	strErrMsg = NULL
	
	'--- �R�}���h���C�������̎擾 ---
	Set objArgs = WScript.Arguments
	
	'--- �p�����[�^���`�F�b�N ---
	If objArgs.Count = 3 Then
		
		'--- �擾�����p�����[�^���� ---
		strTargetDir	= objArgs(0)
		strManipulate	= objArgs(1)
		STR_MANIPULATE_MODE	= objArgs(2)	'�����񑀍샂�[�h
		
		'--- �����J�n���O�o�� ---
	 	Call OutLogMsg(STATUS_TYPE_INFO, MSG_START)

		'������ �p�����[�^�`�F�b�N ������
		'If ChackParam() == False Then
		'End If


		'--- �w��p�X���t�H���_���ǂ����`�F�b�N ---
		If objFS.FolderExists(strTargetDir) Then
			Call CreateFileList(strTargetDir, strManipulate, strErrMsg)
		Else
	 		Call OutLogMsg(STATUS_TYPE_ERROR, ERR_MSG_NOT_FOLDER_PATH)
			'--- �����I�����O�o�� ---
		 	Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
		 	WScript.Quit NG
		End If

	'--- �p�����[�^���s���Ȃ̂ŃG���[�I�� ---
	Else

		'--- �����J�n���O�o�� ---
		Call OutLogMsg(STATUS_TYPE_INFO, MSG_START)
		'--- �G���[���O�o�� ---
		Call OutLogMsg(STATUS_TYPE_ERROR, ERR_MSG_ARG)
		'--- �����I�����O�o�� ---
		Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
		WScript.Quit NG
	End If
	
	'--- �����I�����O�o�� ---
	Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
	
	'--- ����I�� ---
	Set objFS = Nothing
	WScript.Quit OK
End Function

'----------------------------------------------------------------
'�@RenameFileMain�֐�
'----------------------------------------------------------------
'	�p�r�F		�t�@�C�����l�[�������̃��C���֐�
'				���[�h�ɂ��t�@�C�����̐擪�A�������͖����ɕ������t��������
'				�w�蕶������������܂��B
'	�󂯎��l:	[i] strTargetPath	���l�[���Ώۃt�@�C����
'				[i] strPrefix		�w�蕶����
'				[o] strErrMsg		�G���[���b�Z�[�W
'	�߂�l�F	True:����
'				False:�G���[
'----------------------------------------------------------------
Function RenameFileMain(strTargetPath, strPrefix, strErrMsg)
	On Error Resume Next
	
	Dim strLine				'csv�t�@�C������ǂݍ��񂾁A1�s���̕�����
	Dim strRenamedName		'���l�[����̃t�@�C����
	Dim intRenameStatus		'���l�[�������̌���
	
	'--- ������ ---
	RenameFileMain = False
	strErrMsg = NULL
	Err.Number = 0
	
	'--- ���l�[����̃t�@�C�����쐬 ---
	Call MakeRenamedName(strTargetPath, strPrefix, strRenamedName)

	'--- �t�@�C���̃��l�[�� --
	intRenameStatus = RenameFile(strTargetPath, strRenamedName, strPrefix, strErrMsg)
		If intRenameStatus = -1 Then	'--- �G���[ ---
			Exit Function
		ElseIf intRenameStatus = 1 Then	'--- �x�� ---
			Call OutLogMsg(STATUS_TYPE_WARNING, strErrMsg)
		End If
	
	'--- ����I�� ---
	RenameFileMain = True

End Function

'----------------------------------------------------------------
'�@MakeRenamedName�֐�
'----------------------------------------------------------------
'	�p�r�F		���l�[����̃t�@�C�������쐬���܂��B
'	�󂯎��l�F[i] strTargetPath	���l�[���Ώۃt�@�C���p�X
'				[i] strPrefix		�w�蕶����
'				[o] strRenamedName	���l�[����̃t�@�C����
'	�߂�l�F	�Ȃ�
'----------------------------------------------------------------
Sub MakeRenamedName(strTargetPath, strPrefix, strRenamedName)
	Dim strFileName		'�p�X����擾�����t�@�C����

	'--- ���l�[���Ώۃt�@�C���p�X����t�@�C�������擾 ---
	strFileName = objFS.getFileName(strTargetPath)

	'--- ���l�[����̃t�@�C�������쐬
	If STR_MANIPULATE_MODE = STR_MODE_FROMT Then
		strRenamedName = strPrefix & strFileName
	ElseIf STR_MANIPULATE_MODE = STR_MODE_REAR Then
		strRenamedName = strFileName & strPrefix
	ElseIf STR_MANIPULATE_MODE = STR_MODE_DELETE Then
		strRenamedName = Replace(strFileName, strPrefix, "")
	End If
End Sub

'----------------------------------------------------------------
'�@RenameFile�֐�
'----------------------------------------------------------------
'	�p�r�F		�Ώۃt�@�C���������l�[�����܂��B
'	�󂯎��l�F[o] strTargetPath	���l�[���Ώۃt�@�C���̃t���p�X
'				[o] strRenamedName	���l�[����̃t�@�C����
'				[i] strPrefix		�w�蕶����
'				[o] strErrMsg		�G���[���b�Z�[�W
'	�߂�l�F	0:����
'				1:�x��
'				-1:�G���[
'	�G���[�F	���l�[���Ɏ��s�����ꍇ
'	�x���F		���ɓ����̃t�@�C�������݂���ꍇ
'				���l�[���Ώۃt�@�C���p�X����w��̏ꍇ
'				���l�[���Ώۃt�@�C�������݂��Ȃ��ꍇ
'				���l�[���Ώۂ��T�[�o�t�@�C���̃p�X�ł͂Ȃ��ꍇ
'----------------------------------------------------------------
Function RenameFile(strTargetPath, strRenamedName, strPrefix, strErrMsg)
	On Error Resume Next
	Dim objFile

	'--- ������ ---
	RenameFile = -1

	Set objFile = objFS.GetFile(strTargetPath)
	
	'--- ���l�[������ ---
	objFile.name = strRenamedName
	If Err.Number = 70 Then	'--- �������݋֎~�̏ꍇ�̓G���[ ---
		strErrMsg = ERR_MSG_FILE_RENAME & STR_LEFT_SQUARE_BRACKET & strTargetPath & STR_RIGHT_SQUARE_BRACKET
		Exit Function
	ElseIf Err.number = 58 Then		'--- ���ɓ����̃t�@�C�������݂���ꍇ�͌x�� ---
		strErrMsg = WARNING_MSG_FILE_SAME_NAME & STR_LEFT_SQUARE_BRACKET & strRenamedName & STR_RIGHT_SQUARE_BRACKET
		RenameFile = 1
		Exit Function
	ElseIf Err.number = 5 Then		'--- �t�@�C�����Ɏg�p�ł��Ȃ��������w�肳�ꂽ�ꍇ�̓G���[ ---
		strErrMsg = ERR_MSG_CANNOT_FILE_NAME & STR_LEFT_SQUARE_BRACKET & strPrefix & STR_RIGHT_SQUARE_BRACKET
		Exit Function
	End If

	RenameFile = 0
End Function

'----------------------------------------------------------------
'�@OutLogMsg�֐�
'----------------------------------------------------------------
'	�p�r�F		�������O���o�͂��܂��B
'				�w�胍�O�t�@�C���̃I�[�v�����s�����ꍇ�́A�f�t�H���g���O�ɏo�͂��s���܂��B
'				�f�t�H���g���O���o�͂ł��Ȃ������ꍇ�ɂ͕W���o�͂ɏo�͂��܂��B
'	�󂯎��l�F[i] strStatus	�G���[�X�e�[�^�X
'				[i] strErrMsg	�G���[���b�Z�[�W
'	�߂�l�F	�Ȃ�
'	�y�t�H�[�}�b�g�FCSV�`���z
'		�X�e�[�^�X,��������(YYYY/MM/DD hh:mm:ss),���b�Z�[�W(�G���[���e��)
'
'	�y�G���[�X�e�[�^�X�̎�ށz
'		Info�@�@���
'		Warning �x��
'		Error   �G���[
'----------------------------------------------------------------
Sub OutLogMsg(strStatus, strErrMsg)
	On Error Resume Next
	Dim objLog			'���O�t�@�C���I�u�W�F�N�g
	Dim strCurrentDir	'�J�����g�f�B���N�g��
	Dim strLogPath		'���O�t�@�C���p�X

	'--- ������ ---
	Err.Number = 0

	strCurrentDir = objFS.GetParentFolderName(WScript.ScriptFullName)
	strLogPath = strCurrentDir & STR_YEN & STR_LOG_FILE_NAME

	'--- �o�̓t�@�C���I�[�v��(�I�u�W�F�N�g����) ---
	Set objLog = objFS.OpenTextFile(strLogPath, 8, True)
	
	'--- �w�胍�O�t�@�C���̃I�[�v���Ɏ��s�����ꍇ�́A�W���o�� ---
	If Err.Number <> 0 Then
			'--- �G���[��W���o�͂ŏo�͂��ď����𒆎~���� ---
			WScript.Echo STR_DOUBLEQUOT & FormatDateTime(Now) & STR_JOIN_LOG & strStatus  & STR_JOIN_LOG & strErrMsg & STR_DOUBLEQUOT
			WScript.Echo ERR_MSG_FILE_NOT_OPEN & STR_LEFT_SQUARE_BRACKET & strLogPath & STR_RIGHT_SQUARE_BRACKET
		 	WScript.Quit NG
	Else
		'--- �X�e�[�^�X,��������,���b�Z�[�W���o�� ---
		objLog.WriteLine STR_DOUBLEQUOT & FormatDateTime(Now) & STR_JOIN_LOG & strStatus  & STR_JOIN_LOG & strErrMsg & STR_DOUBLEQUOT
	End If

	'--- �N���[�Y ---
	objLog.Close
End Sub

'-----------------------------------------------
'���X�g�쐬���C��
'-----------------------------------------------
Sub CreateFileList(inFolderName, strManipulate, strErrMsg)
	Dim fsoFolder
	Dim fsoSubFolder
	Dim fsoFile
	
	'--- �t�H���_�I�u�W�F�N�g�擾 ---
	Set fsoFolder = objFS.GetFolder(inFolderName)
	
	'--- �t�H���_��/�t�@�C�����[�v ---
	For Each fsoFile In fsoFolder.Files
		'--- ���l�[�� ---
		If  RenameFileMain(fsoFile.Path, strManipulate, strErrMsg) = False Then
			Call OutLogMsg(STATUS_TYPE_ERROR, strErrMsg)
		End If
	Next
	
	'--- �t�H���_��/�T�u�t�H���_���[�v(�T�u�t�H���_���s�v�Ȃ�A���̃��[�v�͕s�v) ---
	For Each fsoSubFolder In fsoFolder.SubFolders
		'--- �T�u�t�H���_�ōċA ---
		Call CreateFileList(fsoSubFolder)
	Next
End Sub