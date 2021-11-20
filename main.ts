// スプレッドシート 時間外勤務表の管理
const ssId = '1EhXDA7gHDe7hNZqbjDhERZabgYCrSsvIXaL-ohRr2Eg'
const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ssId)
const sheet: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName('今月')
const sheetId = ss.getActiveSheet().getSheetId()
const scriptProperties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties()

// トリガーで毎日AM6:00に実行する
function main() {
	console.log('Main function !')

	// 1日かどうかチェックする．1日じゃなければ処理終了する．
	if (!isCheckDateOne_()) return

	// シートをコピー
	copySheet_()
	const fileName = createFileName_()
	const pdfBlob = createPdfBlob_(ss, fileName)
	const pdfFile = createPdfFile_(pdfBlob)
	const fileUrl = getFileUrl_(pdfFile)

	/* Version2
	   Urlfetchを使用したバージョン
	 */
	const { pdfBlob_ver2, file_ver2 } = createPdf_ver2_(FOLDER_ID, ssId, sheetId, fileName)
	const fileUrl_ver2 = getFileUrl_(file_ver2)

	// メールに送信
	sendEmail_(pdfBlob, fileName, fileUrl)
	sendEmail_(pdfBlob_ver2, fileName, fileUrl_ver2)

	// LINEに送信
	const message = `\n\n今月の時間外勤務表を送信しました．\n\n${fileName}\n${fileUrl}`
	sendLineNotify(message)
}

function sendLineNotify(message: string = createFileName_()) {
	const LINE_NOTIFY_API_TOKEN = scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN')
	const LINE_NOTIFY_API_URL = 'https://notify-api.line.me/api/notify'

	const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
		method: 'post',
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
			Authorization: `Bearer ${LINE_NOTIFY_API_TOKEN}`,
		},
		payload: { message },
	}

	// LINEに通知
	UrlFetchApp.fetch(LINE_NOTIFY_API_URL, options)
}

// POSTリクエストが来たときの処理 ---------------------------------------------------
function doPost(e: GoogleAppsScript.Events.DoPost) {
	// NextAppからPOSTされたデータを取得する
	const data: OvertimePostData = JSON.parse(e.postData.contents).data
	const { radiologist, modality, date, start, end, description } = data

	// タイムスタンプ生成
	const created_at = new Date()

	//! スプレッドシートにフォームデータを書き込んで、lineNotify関数を実行する．
	// スプレッドシートに追加する項目↓
	// [uuid, created_at, created_at, date, radiologist, modality, start, end, description]
	const uuid = Utilities.getUuid()

	// date,start,endはstring型なのでDate型に変換する．
	const parsed_date = new Date(date)
	const parsed_start = toDateFromStrTime_(start)
	const parsed_end = toDateFromStrTime_(end)
	const diff_time = ((parsed_end.getTime() - parsed_start.getTime()) / 1000 / 60).toFixed(1) // 分
	const diff_time_2 = ((parsed_end.getTime() - parsed_start.getTime()) / 1000 / 60 / 60).toFixed(1) // 時間

	// 10列目に残業時間（分）のヘッダーが無ければ入力する．
	const Header_10 = '残業時間（分）'
	const diff_time_cell = sheet?.getRange(1, 10)
	if (diff_time_cell?.getValue() !== Header_10) {
		diff_time_cell?.setValue(Header_10)
	}
	// 11列目に残業時間（時間）のヘッダーが無ければ 入力する．
	const Header_11 = '残業時間（時間）'
	const diff_time_cell_2 = sheet?.getRange(1, 11)
	if (diff_time_cell_2?.getValue() !== Header_11) {
		diff_time_cell_2?.setValue(Header_11)
	}
	// 12列目に合計残業時間（時間(分)）のヘッダーが無ければ 入力する．
	const Header_12 = '合計残業時間（時間(分)'
	const total_time_cell = sheet?.getRange(1, 12)
	if (total_time_cell?.getValue() !== Header_12) {
		total_time_cell?.setValue(Header_12)
	}

	// spreadsheetに追加する．
	sheet?.appendRow([
		uuid,
		created_at,
		created_at,
		radiologist,
		modality,
		parsed_date,
		start,
		end,
		description,
		diff_time,
		diff_time_2,
	])

	// 合計時間を計算
	let total_time = 0
	sheet
		?.getDataRange()
		.getValues()
		.forEach((row) => {
			const [
				uuid,
				created_at,
				updated_at,
				radiologist,
				modality,
				parsed_date,
				start,
				end,
				description,
				diff_time,
				diff_time_2,
			] = row

			const _total = Number(diff_time) + Number(diff_time_2)
			total_time = _total
		})

	// 12行目2列目に合計時間を入力
	sheet?.getRange(2, 12, 1, 1).setValue(total_time.toFixed())

	// スプレッドシートに書き込まれるまで少し待機
	Utilities.sleep(5 * 1000)

	// TODO シートネーム、sheetUrl取得
	const ssUrl = ss.getUrl()
	const crr_sheetName = ss.getActiveSheet().getName()

	const postMessage = `「NextAppから${crr_sheetName}」へ時間外勤務を登録しました．\n\n実施日: ${date}\n実施者: ${radiologist}\nモダリティ: ${modality}\n時間: ${start} 〜 ${end}\n業務内容: ${description}\n\n${ssUrl}`
	// LINEへ通知
	sendLineNotify(postMessage)
}

function toDateFromStrTime_(time: string): Date {
	// time = '12:05'
	const d = new Date()
	const [HH, mm] = time.split(':')
	d.setHours(parseInt(HH))
	d.setMinutes(parseInt(mm))
	return d
}

const copySheet_ = () => {
	try {
		// 新シート生成
		// 既存シート数
		const index = ss.getNumSheets()
		// シート名生成
		const fileName = createFileName_()
		// シート挿入
		ss.insertSheet(fileName, index + 1)
		// 新しいシートを作成して旧シートからコピーする
		if (sheet !== null) {
			// 最終行
			const lr = sheet?.getLastRow()
			// 最終列
			const lc = sheet?.getLastColumn()
			// 新シート作成
			const newSheet = ss.getSheetByName(fileName)
			// 旧シートからデータを転記
			newSheet?.getRange(1, 1, lr, lc).setValues(sheet?.getRange(1, 1, lr, lc).getValues())
			// おそらくフォーマットが狂うので整形
			// （ここでは7, 8列目に残業開始時間、終了時間が並んでいるものと想定）
			newSheet?.getRange(2, 7, lr - 1, 2).setNumberFormat('hh:mm')
			//（ここでは2,3,6列目の作成日時、変更日時2箇所）
			newSheet?.getRange(2, 2, lr - 1, 2).setNumberFormat('yyyy-mm-dd')
			// （ここでは6列目の実施日）
			newSheet?.getRange(2, 6, lr - 1, 1).setNumberFormat('yyyy-mm-dd')

			// 旧シート初期化
			if (new Date().getDate() === 1) {
				sheet?.deleteRows(2, lr - 1) // あえて.getRange().clear()は使わない
			}
		}

		// トリガーが失敗したら知らせる
	} catch (e) {
		console.log('コピーシートError:: ', e)
		sendLineNotify(`コピーシートError:: \n\n${e.message}`)
	}
}

function setScriptProperty_() {
	scriptProperties.setProperty('LINE_NOTIFY_API_TOKEN', 'ys9d5Voj5gYDXK2GuAbskcpzq77gt2XDs3vMiKwseQB')

	console.log(scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN') ?? '')
}

function test() {
	console.log('This is Test Script !')

	// シートをコピー
	copySheet_()
	// let fileName = createFileName_()
	const fileName = '_' + Utilities.getUuid()
	const pdfBlob = createPdfBlob_(ss, fileName)
	const pdfFile = createPdfFile_(pdfBlob)
	const fileUrl = getFileUrl_(pdfFile)

	/* Version2
	   Urlfetchを使用したバージョン
	 */
	const { pdfBlob_ver2, file_ver2 } = createPdf_ver2_(FOLDER_ID, ssId, sheetId, fileName)
	const fileUrl_ver2 = getFileUrl_(file_ver2)

	// メールに送信
	sendEmail_(pdfBlob, fileName, fileUrl)
	sendEmail_(pdfBlob_ver2, fileName, fileUrl_ver2)

	// LINEに送信
	const message = `\n\nTestです今月の時間外勤務表を送信しました．.\n\n${fileName}\n${fileUrl}`
	sendLineNotify(message)
}
