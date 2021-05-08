// ファイル名を作成する．'202105'
const createFileName = (): string => {
	// 日付
	const date = new Date()

	// 年
	const yy = date.getFullYear()
	const str_yy = String(yy)

	//! 月（0~11が返ることに注意）
	const mm = date.getMonth()

	let str_mm: string = ''
	// １０月未満の場合は頭に０を付す
	if (mm + 1 < 10) {
		str_mm = '0' + String(mm + 1)
		// 1月の場合は０が返るから
	} else if (mm === 0) {
		str_mm = '01'
	}

	// シート名生成
	const fileName = str_yy + str_mm
	return fileName
}

// スプレッドシートをPDFとして取得する
const createPdfBlob = (
	spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
	fileName: string,
): GoogleAppsScript.Base.Blob => {
	//! スプレッドシートのすべてのシートをPDFとして取得されてしまう
	//todo スプレッドシートを1シートずつPDF化する方法はないのか？ 無さそう．
	const pdfBlob = spreadSheet.getAs('application/pdf').setName(`${fileName}.pdf`)
	// const pdfBlob = spreadSheet.getBlob().getAs('image/jpeg').setName(`${fileName}.jpeg`)
	// -->Exception: application/pdf から image/jpeg への変換はサポートされていません。
	return pdfBlob
}

// googleDriveの指定のフォルダ("過去データ")へ保存する
const createPdfFile = (blob: GoogleAppsScript.Base.Blob): GoogleAppsScript.Drive.File => {
	// Folderオブジェクト.createFile(Blobオブジェクト)
	const FOLDER_ID = '1oVv95yEt3Pm1itm8ocqKW8ij88ZXTobG'
	const folder = DriveApp.getFolderById(FOLDER_ID) //フォルダを指定
	const pdfFile = folder.createFile(blob)
	return pdfFile
}

// googleDriveに保存したファイルのURLを取得する
const getFileUrl = (pdfFile: GoogleAppsScript.Drive.File): string => {
	// Fileオブジェクト.getUrl()
	const fileUrl = pdfFile.getUrl()
	return fileUrl
}

// 本日が1日かどうかを判定する関数:boolean
const isCheckDate = () => {
	const today = new Date().getDate()
	// もし１ならtrueを返す
	return today === 1
}

// 自分宛てにPDFBlobをメールで送信する
const sendEmail = (pdfBlob: GoogleAppsScript.Base.Blob, url: string): void => {
	const MY_ADDRESS = 'kfcxd953pelo@gmail.com'
	const fileUrl = url ? url : ''
	MailApp.sendEmail(
		MY_ADDRESS, // 宛先
		'今月の時間外勤務表', // 件名
		`PDFを送りました\n` + fileUrl, //本文
		// 添付ファイル(pdf)
		{ attachments: [pdfBlob] },
	)
}
