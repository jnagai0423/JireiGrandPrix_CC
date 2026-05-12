/** スプレッドシートを開いたときに手動実行メニューを追加します。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gmail取り込み')
    .addItem('今すぐ取り込み', 'importCgMailToSheet')
    .addToUi();
}
