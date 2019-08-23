using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ID3compare2
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {

        string sPath;
        string[] errorFilesList;
        int sfCount;
        int i;
        TagLib.File oneMP3File;

        private BackgroundWorker BW;

        public MainWindow()
        {
            InitializeComponent();
            FolderNameLabel.Content = "フォルダを選択するかドラッグ&ドロップしてください。";
            ResultTextBox.Text = "";
            MainProgressBar.Visibility = Visibility.Hidden;

        }

        private void ResultTextBox_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string f in dragFilePathArr)
            {
                if (Directory.Exists(f))
                {
                    sPath = dragFilePathArr[0];
                    sfCount = Directory.GetFiles(sPath, "*.mp3", SearchOption.TopDirectoryOnly).Length;

                    FolderNameLabel.Content = "Path : " + sPath + " / MP3 : " + sfCount + "files";

                    // 初期化
                    ResultTextBox.Clear();
                    errorFilesList = new string[1];
                    BW = new BackgroundWorker();

                    /* ボタンの名前を変更 */
                    SelectFolderButton.Content = "中止";

                    /* プログレスバーを設定 */
                    MainProgressBar.Maximum = sfCount;
                    MainProgressBar.Visibility = Visibility.Visible;

                    /* BackgroundWorkerのProgressChangedイベントが発生するようにする */
                    BW.WorkerReportsProgress = true;

                    // 途中で中止できるようにする
                    BW.WorkerSupportsCancellation = true;

                    // イベントハンドラを追加
                    BW.DoWork += new DoWorkEventHandler(DoWork);
                    BW.ProgressChanged += new ProgressChangedEventHandler(BW_ProgressChanged);
                    BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_RunWorkerCompleted);

                    /* DoWorkで取得できるパラメータを指定して、処理を開始する
                       パラメータが必要なければ省略できる */
                    BW.RunWorkerAsync();
                }
            }
        }


        private void ResultTextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            var fileList = ((DataObject)e.Data).GetFileDropList();

            if (fileList.Count > 0)
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
            }
        }



        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            BW = new BackgroundWorker();

            // ボタンの名前で判断
            if (SelectFolderButton.Content is "選択")
            {
#pragma warning disable IDE0068 // 推奨される dispose パターンを使用する
                var dialog = new CommonOpenFileDialog("フォルダ選択")
                {

                    // フォルダ選択モード
                    IsFolderPicker = true,
                    Multiselect = false
                };
#pragma warning restore IDE0068 // 推奨される dispose パターンを使用する

                //フォルダを選択するダイアログを表示する
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    sPath = dialog.FileName;
                    sfCount = Directory.GetFiles(sPath, "*.mp3", SearchOption.TopDirectoryOnly).Length;

                    FolderNameLabel.Content = "Path : " + sPath + " / MP3 : " + sfCount + "files";

                    // 初期化
                    ResultTextBox.Clear();
                    errorFilesList = new string[1];

                    /* ボタンの名前を変更 */
                    SelectFolderButton.Content = "中止";

                    /* プログレスバーを設定 */
                    MainProgressBar.Maximum = sfCount;
                    MainProgressBar.Visibility = Visibility.Visible;

                    /* BackgroundWorkerのProgressChangedイベントが発生するようにする */
                    BW.WorkerReportsProgress = true;

                    // 途中で中止できるようにする
                    BW.WorkerSupportsCancellation = true;

                    // イベントハンドラを追加
                    BW.DoWork += new DoWorkEventHandler(DoWork);
                    BW.ProgressChanged += new ProgressChangedEventHandler(BW_ProgressChanged);
                    BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_RunWorkerCompleted);

                    /* DoWorkで取得できるパラメータを指定して、処理を開始する
                       パラメータが必要なければ省略できる */
                    BW.RunWorkerAsync();
                }
                else
                {
                    Console.WriteLine("キャンセルされました");
                }
            }
            else
            {
                /* キャンセルする */
                BW.CancelAsync();
            }
        }



        // 別スレッドで回す重たい処理（つまりはメインのループ）
        private void DoWork(object sender, DoWorkEventArgs e)
        {
            i = 0;
            string oneMP3FileComposers;
  
            // 比較元フォルダのファイル名を配列に入れる
            string[] filenameArray = Directory.GetFiles(sPath, "*.mp3", SearchOption.TopDirectoryOnly);

            // その配列をforeachに投げ込んで処理
            foreach (string filename in filenameArray)
            {
                // キャンセルされたかどうかチェック
                if (BW.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // パスを除いたファイル名と、ID3タグの作曲者を取得
                    // 【20190819】taglibを利用して、作曲者名を数字じゃなく要素として取得してみよう計画！
                    string oneFileName = System.IO.Path.GetFileName(filename);
                    oneMP3File = TagLib.File.Create(filename);

                    /* デバッグ用
                    if (oneMP3File.Tag.Composers.Length > 0)
                    {
                        Console.WriteLine(oneMP3File.Tag.Composers.Length);
                    }
                    */

                    try
                    {
                        oneMP3FileComposers = oneMP3File.Tag.Composers[0];
                    }
                    catch (IndexOutOfRangeException)
                    {
                        oneMP3FileComposers = "なし";
                    }
                    
                    // ファイル名（拡張子をmp2にして合わせたもの）と作曲者名部分を比較
                    if (oneFileName.Replace("mp3","mp2") != oneMP3FileComposers)
                    {
                        // 見つかったファイルを配列に入れておく
                        if (errorFilesList.Length > 1)
                        {
                            errorFilesList[errorFilesList.Length -1] = oneFileName + "（タグ：" + oneMP3FileComposers + "）";
                            Array.Resize(ref errorFilesList, errorFilesList.Length + 1);
                        }
                        else
                        {
                            errorFilesList[0] = oneFileName + "（タグ：" + oneMP3FileComposers + "）";
                            Array.Resize(ref errorFilesList, errorFilesList.Length + 1);
                        }
                        //Console.WriteLine("ファイル名: " + oneFileName);
                    }

                    i++;
                    // BackgroundWorkerに現状を送る
                    BW.ReportProgress(i);
                }
            }
        }


        // BackgroundWorkerのProgressChangedイベントハンドラ
        // コントロールの操作は必ずここで行い、DoWorkでは絶対にしない
        private void BW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // ProgressBarの値を変更する
            MainProgressBar.Value = e.ProgressPercentage;
        }
        
        // BackgroundWorkerのRunWorkerCompletedイベントハンドラ
        // 処理が終わったときに呼び出される
        private void BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((e.Cancelled == true))
            {
                ResultTextBox.Text = "キャンセルされました。";
            }
            else if (!(e.Error == null))
            {
                ResultTextBox.Text = "エラー：" + e.Error.Message;
            }
            else
            {
                ResultTextBox.Text = "完了しました。" + Environment.NewLine;
                
                if (errorFilesList.Length > 1) // ID3Tagがおかしいファイルがあったら
                {
                    ResultTextBox.AppendText("以下のファイルのID3タグが正しく書き込めていないようです。");

                    foreach(string eft in errorFilesList)
                    {
                        ResultTextBox.AppendText(Environment.NewLine + eft);
                    }
                }
                else
                {
                    ResultTextBox.Text += "全ファイルのID3タグが正しく書き込まれています。";
                }
            }

            // 各種初期化
            errorFilesList.Initialize();
            MainProgressBar.Visibility = Visibility.Hidden;
            SelectFolderButton.Content = "選択";
            BW.Dispose();
        }

    }
}
