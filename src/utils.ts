// ファイル名を作成する．'202105' スクリプト実行の前月を取得したい
const createFileName_ = (): string => {
	const d = new Date()
	const lastMonth = d.getMonth() - 1
	d.setMonth(lastMonth)
	return Utilities.formatDate(d, 'JST', 'yyyyMM')
}

// スプレッドシートをPDFとして取得する
const createPdfBlob_ = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, fileName: string): GoogleAppsScript.Base.Blob => {
	const pdfBlob = ss.getAs('application/pdf').setName(`${fileName}.pdf`)
	return pdfBlob
}

// googleDriveの指定のフォルダ("過去データ")へ保存する
//! すべてのsheetがPDFに変換されたしまう．
const FOLDER_ID = '1oVv95yEt3Pm1itm8ocqKW8ij88ZXTobG'
const createPdfFile_ = (blob: GoogleAppsScript.Base.Blob): GoogleAppsScript.Drive.File => {
	const folder = DriveApp.getFolderById(FOLDER_ID) //フォルダを指定
	const pdfFile = folder.createFile(blob)
	return pdfFile
}

// googleDriveに保存したファイルのURLを取得する
const getFileUrl_ = (file: GoogleAppsScript.Drive.File): string => {
	const fileUrl = file.getUrl()
	return fileUrl
}

// 本日が1日かどうかを判定する関数:boolean
const isCheckDateOne_ = () => {
	const today = new Date().getDate()
	return today === 1
}

// 自分宛てにPDFBlobをメールで送信する
const sendEmail_ = (pdfBlob: GoogleAppsScript.Base.Blob, fileName: string, fileUrl: string): void => {
	const MY_ADDRESS = 'kfcxd953pelo@gmail.com'
	MailApp.sendEmail(
		MY_ADDRESS, // 宛先
		`今月の時間外勤務表_${fileName}`, // 件名
		`PDFを送りました\n` + fileUrl, // 本文
		{ attachments: [pdfBlob] }, // 添付ファイル(pdf)
	)
}

//PDFを作成し指定したフォルダーに保存する関数
//以下4つの引数を指定する必要がある
//1: フォルダーID (FOLDER_ID)
//2: スプレッドシートID (sheetId)
//3: シートID (shId)
//4: ファイル名 (fileName)
//! 一つのsheetのみをPDFに変換できる．
function createPdf_ver2_(folderId: string, ssId: string, sheetId: number, fileName: string) {
	//PDFを作成するためのベースとなるURL
	const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?gid=' + sheetId

	//★★★自由にカスタマイズしてください★★★
	//PDFのオプションを指定
	const pdfOptions =
		'&exportFormat=pdf&format=pdf' +
		'&size=A4' + //用紙サイズ (A4)
		'&portrait=true' + //用紙の向き true: 縦向き / false: 横向き
		'&fitw=true' + //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
		'&top_margin=0.50' + //上の余白
		'&right_margin=0.50' + //右の余白
		'&bottom_margin=0.50' + //下の余白
		'&left_margin=0.50' + //左の余白
		'&horizontal_alignment=CENTER' + //水平方向の位置
		'&vertical_alignment=TOP' + //垂直方向の位置
		'&printtitle=false' + //スプレッドシート名の表示有無
		'&sheetnames=true' + //シート名の表示有無
		'&gridlines=false' + //グリッドラインの表示有無
		'&fzr=false' + //固定行の表示有無
		'&fzc=false' //固定列の表示有無;

	//PDFを作成するためのURL
	const url = baseUrl + pdfOptions

	//アクセストークンを取得する
	const token = ScriptApp.getOAuthToken()

	//headersにアクセストークンを格納する
	const options = {
		headers: {
			Authorization: 'Bearer ' + token,
		},
		muteHttpExceptions: true,
	}

	//PDFを作成する
	const pdfBlob_ver2 = UrlFetchApp.fetch(url, options)
		.getBlob()
		.setName(fileName + '_ver2' + '.pdf')

	//PDFの保存先フォルダー
	//フォルダーIDは引数のfolderIdを使用します
	const folder = DriveApp.getFolderById(folderId)

	//PDFを指定したフォルダに保存する
	const file_ver2 = folder.createFile(pdfBlob_ver2)

	return { pdfBlob_ver2, file_ver2 }
}
