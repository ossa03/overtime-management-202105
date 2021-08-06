// ファイル名を作成する．'202105'
const createFileName_ = (): string => {
	return Utilities.formatDate(new Date(), 'JST', 'yyyyMM')
}

// スプレッドシートをPDFとして取得する
const createPdfBlob_ = (
	spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
	fileName: string,
): GoogleAppsScript.Base.Blob => {
	const pdfBlob = spreadSheet.getAs('application/pdf').setName(`${fileName}.pdf`)
	return pdfBlob
}

// googleDriveの指定のフォルダ("過去データ")へ保存する
const createPdfFile_ = (blob: GoogleAppsScript.Base.Blob): GoogleAppsScript.Drive.File => {
	const FOLDER_ID = '1oVv95yEt3Pm1itm8ocqKW8ij88ZXTobG'
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
