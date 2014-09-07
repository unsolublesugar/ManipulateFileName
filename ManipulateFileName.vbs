Option Explicit
'================================================================
'ファイル名操作スクリプト ManipulateFileName.vbs
'================================================================
'	指定フォルダ配下のファイルすべてに対して、ファイル名の文字列を一括で変更します。
'	モード指定により、指定文字列をファイル名の先頭もしくは末尾に付加したり
'	ファイル名から指定文字列を除去することも出来ます。
'
'【引数】
'	変換対象ファイルのあるフォルダパス
'	付加したいor除去したい文字列
'	モード
'		1:先頭に文字列付加
'		2:末尾に文字列付加
'		3:指定文字列を削除)
'
'【戻り値】
'   OK  0	正常終了
'   NG  1	異常終了

'【ログ出力】
'	処理実行、終了時と、エラーおよび警告が発生した場合は、実行フォルダにログを出力します。
'
'	【フォーマット：CSV形式】
'	"処理日時(YYYY/MM/DD hh:mm:ss)","ステータス","メッセージ"
'
'	【ステータスの種類】
'		Info	(情報)
'		Warning	(警告)
'		Error	(エラー)
'
'
'----------------------------------------------------------------
'　定数
'----------------------------------------------------------------

	'処理メッセージ
	Const MSG_START		=	"処理を開始します。"
	Const MSG_END		=	"処理を終了しました。"

	'エラーメッセージ
	Const ERR_MSG_ARG						=	"不正なパラメータが指定されました。"
	Const ERR_MSG_NOT_FOLDER_PATH			=	"指定されたパスはフォルダではありません。"
	Const ERR_MSG_FILE_NOT_OPENN			=	"ファイルオープンに失敗しました。"
	Const ERR_MSG_CANNOT_FILE_NAME			=	"ファイル名に使用できない文字が指定されました。"

	'警告メッセージ
	Const WARNING_MSG_FILE_SAME_NAME			=	"同一ファイル名のファイルが既に存在します。"

	'ステータスの種類
	Const STATUS_TYPE_INFO		=	"Info"
	Const STATUS_TYPE_WARNING	=	"Warning"
	Const STATUS_TYPE_ERROR		=	"Error"

	'記号
	Const STR_YEN					=	"\"
	Const STR_DOUBLEQUOT			=	""""
	Const STR_LEFT_SQUARE_BRACKET	=	"["
	Const STR_RIGHT_SQUARE_BRACKET	=	"]"
	Const STR_JOIN_LOG				= 	""","""
	
	'デフォルトログ出力パス
	Const STR_LOG_FILE_NAME		=	"ManipulateFileName.log"
	
	'文字列操作モード
	Const STR_MODE_FROMT	=	"1"		'先頭に付加
	Const STR_MODE_REAR		=	"2"		'末尾に付加
	Const STR_MODE_DELETE	=	"3"		'削除
	
'----------------------------------------------------------------
'　ステータス定数宣言
'----------------------------------------------------------------
	Const OK	= "0"	'正常
	Const NG	= "1"	'異常
'----------------------------------------------------------------

Dim objFS
Dim STR_MANIPULATE_MODE		'文字列操作モード

'--- FileSystemObjectの生成 ---
Set objFS = CreateObject("Scripting.FileSystemObject")
STR_MANIPULATE_MODE	=	0


Call Main()



'----------------------------------------------------------------
'メイン処理を行います。
'
'@param		無し
'@return	RETURN_TRUE		正常
'			RETURN_FALSE	エラーor警告
'----------------------------------------------------------------
Function Main()
	Dim objArgs			'引数取得用オブジェクト
	Dim strErrMsg		'エラーメッセージ
	Dim intMode			'文字列操作モード
	Dim strTargetDir	'対象ディレクトリ
	Dim strManipulate	'操作する文字列
	Dim strLogPath		'ログファイルパス
	Dim objParam

	'--- 初期化 ---
	strErrMsg = NULL
	
	'--- コマンドライン引数の取得 ---
	Set objArgs = WScript.Arguments
	
	'--- パラメータ数チェック ---
	If objArgs.Count = 3 Then
		
		'--- 取得したパラメータを代入 ---
		strTargetDir	= objArgs(0)
		strManipulate	= objArgs(1)
		STR_MANIPULATE_MODE	= objArgs(2)	'文字列操作モード
		
		'--- 処理開始ログ出力 ---
	 	Call OutLogMsg(STATUS_TYPE_INFO, MSG_START)

		'◆◆◆ パラメータチェック ◆◆◆
		'If ChackParam() == False Then
		'End If


		'--- 指定パスがフォルダかどうかチェック ---
		If objFS.FolderExists(strTargetDir) Then
			Call CreateFileList(strTargetDir, strManipulate, strErrMsg)
		Else
	 		Call OutLogMsg(STATUS_TYPE_ERROR, ERR_MSG_NOT_FOLDER_PATH)
			'--- 処理終了ログ出力 ---
		 	Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
		 	WScript.Quit NG
		End If

	'--- パラメータが不正なのでエラー終了 ---
	Else

		'--- 処理開始ログ出力 ---
		Call OutLogMsg(STATUS_TYPE_INFO, MSG_START)
		'--- エラーログ出力 ---
		Call OutLogMsg(STATUS_TYPE_ERROR, ERR_MSG_ARG)
		'--- 処理終了ログ出力 ---
		Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
		WScript.Quit NG
	End If
	
	'--- 処理終了ログ出力 ---
	Call OutLogMsg(STATUS_TYPE_INFO, MSG_END)
	
	'--- 正常終了 ---
	Set objFS = Nothing
	WScript.Quit OK
End Function

'----------------------------------------------------------------
'　RenameFileMain関数
'----------------------------------------------------------------
'	用途：		ファイルリネーム処理のメイン関数
'				モードによりファイル名の先頭、もしくは末尾に文字列を付加したり
'				指定文字列を除去します。
'	受け取る値:	[i] strTargetPath	リネーム対象ファイル名
'				[i] strPrefix		指定文字列
'				[o] strErrMsg		エラーメッセージ
'	戻り値：	True:正常
'				False:エラー
'----------------------------------------------------------------
Function RenameFileMain(strTargetPath, strPrefix, strErrMsg)
	On Error Resume Next
	
	Dim strLine				'csvファイルから読み込んだ、1行分の文字列
	Dim strRenamedName		'リネーム後のファイル名
	Dim intRenameStatus		'リネーム処理の結果
	
	'--- 初期化 ---
	RenameFileMain = False
	strErrMsg = NULL
	Err.Number = 0
	
	'--- リネーム後のファイル名作成 ---
	Call MakeRenamedName(strTargetPath, strPrefix, strRenamedName)

	'--- ファイルのリネーム --
	intRenameStatus = RenameFile(strTargetPath, strRenamedName, strPrefix, strErrMsg)
		If intRenameStatus = -1 Then	'--- エラー ---
			Exit Function
		ElseIf intRenameStatus = 1 Then	'--- 警告 ---
			Call OutLogMsg(STATUS_TYPE_WARNING, strErrMsg)
		End If
	
	'--- 正常終了 ---
	RenameFileMain = True

End Function

'----------------------------------------------------------------
'　MakeRenamedName関数
'----------------------------------------------------------------
'	用途：		リネーム後のファイル名を作成します。
'	受け取る値：[i] strTargetPath	リネーム対象ファイルパス
'				[i] strPrefix		指定文字列
'				[o] strRenamedName	リネーム後のファイル名
'	戻り値：	なし
'----------------------------------------------------------------
Sub MakeRenamedName(strTargetPath, strPrefix, strRenamedName)
	Dim strFileName		'パスから取得したファイル名

	'--- リネーム対象ファイルパスからファイル名を取得 ---
	strFileName = objFS.getFileName(strTargetPath)

	'--- リネーム後のファイル名を作成
	If STR_MANIPULATE_MODE = STR_MODE_FROMT Then
		strRenamedName = strPrefix & strFileName
	ElseIf STR_MANIPULATE_MODE = STR_MODE_REAR Then
		strRenamedName = strFileName & strPrefix
	ElseIf STR_MANIPULATE_MODE = STR_MODE_DELETE Then
		strRenamedName = Replace(strFileName, strPrefix, "")
	End If
End Sub

'----------------------------------------------------------------
'　RenameFile関数
'----------------------------------------------------------------
'	用途：		対象ファイル名をリネームします。
'	受け取る値：[o] strTargetPath	リネーム対象ファイルのフルパス
'				[o] strRenamedName	リネーム後のファイル名
'				[i] strPrefix		指定文字列
'				[o] strErrMsg		エラーメッセージ
'	戻り値：	0:正常
'				1:警告
'				-1:エラー
'	エラー：	リネームに失敗した場合
'	警告：		既に同名のファイルが存在する場合
'				リネーム対象ファイルパスが空指定の場合
'				リネーム対象ファイルが存在しない場合
'				リネーム対象がサーバファイルのパスではない場合
'----------------------------------------------------------------
Function RenameFile(strTargetPath, strRenamedName, strPrefix, strErrMsg)
	On Error Resume Next
	Dim objFile

	'--- 初期化 ---
	RenameFile = -1

	Set objFile = objFS.GetFile(strTargetPath)
	
	'--- リネーム処理 ---
	objFile.name = strRenamedName
	If Err.Number = 70 Then	'--- 書き込み禁止の場合はエラー ---
		strErrMsg = ERR_MSG_FILE_RENAME & STR_LEFT_SQUARE_BRACKET & strTargetPath & STR_RIGHT_SQUARE_BRACKET
		Exit Function
	ElseIf Err.number = 58 Then		'--- 既に同名のファイルが存在する場合は警告 ---
		strErrMsg = WARNING_MSG_FILE_SAME_NAME & STR_LEFT_SQUARE_BRACKET & strRenamedName & STR_RIGHT_SQUARE_BRACKET
		RenameFile = 1
		Exit Function
	ElseIf Err.number = 5 Then		'--- ファイル名に使用できない文字が指定された場合はエラー ---
		strErrMsg = ERR_MSG_CANNOT_FILE_NAME & STR_LEFT_SQUARE_BRACKET & strPrefix & STR_RIGHT_SQUARE_BRACKET
		Exit Function
	End If

	RenameFile = 0
End Function

'----------------------------------------------------------------
'　OutLogMsg関数
'----------------------------------------------------------------
'	用途：		処理ログを出力します。
'				指定ログファイルのオープン失敗した場合は、デフォルトログに出力を行います。
'				デフォルトログが出力できなかった場合には標準出力に出力します。
'	受け取る値：[i] strStatus	エラーステータス
'				[i] strErrMsg	エラーメッセージ
'	戻り値：	なし
'	【フォーマット：CSV形式】
'		ステータス,処理日時(YYYY/MM/DD hh:mm:ss),メッセージ(エラー内容等)
'
'	【エラーステータスの種類】
'		Info　　情報
'		Warning 警告
'		Error   エラー
'----------------------------------------------------------------
Sub OutLogMsg(strStatus, strErrMsg)
	On Error Resume Next
	Dim objLog			'ログファイルオブジェクト
	Dim strCurrentDir	'カレントディレクトリ
	Dim strLogPath		'ログファイルパス

	'--- 初期化 ---
	Err.Number = 0

	strCurrentDir = objFS.GetParentFolderName(WScript.ScriptFullName)
	strLogPath = strCurrentDir & STR_YEN & STR_LOG_FILE_NAME

	'--- 出力ファイルオープン(オブジェクト生成) ---
	Set objLog = objFS.OpenTextFile(strLogPath, 8, True)
	
	'--- 指定ログファイルのオープンに失敗した場合は、標準出力 ---
	If Err.Number <> 0 Then
			'--- エラーを標準出力で出力して処理を中止する ---
			WScript.Echo STR_DOUBLEQUOT & FormatDateTime(Now) & STR_JOIN_LOG & strStatus  & STR_JOIN_LOG & strErrMsg & STR_DOUBLEQUOT
			WScript.Echo ERR_MSG_FILE_NOT_OPEN & STR_LEFT_SQUARE_BRACKET & strLogPath & STR_RIGHT_SQUARE_BRACKET
		 	WScript.Quit NG
	Else
		'--- ステータス,処理日時,メッセージを出力 ---
		objLog.WriteLine STR_DOUBLEQUOT & FormatDateTime(Now) & STR_JOIN_LOG & strStatus  & STR_JOIN_LOG & strErrMsg & STR_DOUBLEQUOT
	End If

	'--- クローズ ---
	objLog.Close
End Sub

'-----------------------------------------------
'リスト作成メイン
'-----------------------------------------------
Sub CreateFileList(inFolderName, strManipulate, strErrMsg)
	Dim fsoFolder
	Dim fsoSubFolder
	Dim fsoFile
	
	'--- フォルダオブジェクト取得 ---
	Set fsoFolder = objFS.GetFolder(inFolderName)
	
	'--- フォルダ内/ファイルループ ---
	For Each fsoFile In fsoFolder.Files
		'--- リネーム ---
		If  RenameFileMain(fsoFile.Path, strManipulate, strErrMsg) = False Then
			Call OutLogMsg(STATUS_TYPE_ERROR, strErrMsg)
		End If
	Next
	
	'--- フォルダ内/サブフォルダループ(サブフォルダが不要なら、このループは不要) ---
	For Each fsoSubFolder In fsoFolder.SubFolders
		'--- サブフォルダで再帰 ---
		Call CreateFileList(fsoSubFolder)
	Next
End Sub