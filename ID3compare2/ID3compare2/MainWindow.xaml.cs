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
        string processFile;
        int sfCount;
        int i;

        private BackgroundWorker bw = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();
            FolderNameLabel.Content = "フォルダを選択してください";
            CompareButton.IsEnabled = false;
            ResultTextBlock.Text = "";
        }

       
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog("フォルダ選択");

            // フォルダ選択モード
            dialog.IsFolderPicker = true;
            dialog.Multiselect = false;

            //フォルダを選択するダイアログを表示する
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                sPath = dialog.FileName;

                sfCount = System.IO.Directory.GetFiles(sPath, "*.mp3", SearchOption.TopDirectoryOnly).Length;

                FolderNameLabel.Content = "Path : " + sPath + " / MP3 : " + sfCount + "files";
                CompareButton.IsEnabled = true;
                ResultTextBlock.Text = "ここに照合結果が表示されます";

            }
            else
            {
                Console.WriteLine("キャンセルされました");
            }
                        
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (bw.IsBusy != true)
            {
                /* 実行ボタンの名前を変更 */
                CompareButton.Content = "Abort";
  
                /* プログレスバーの最大値を設定 */
                MainProgressBar.Maximum = sfCount;

                /* BackgroundWorkerのProgressChangedイベントが発生するようにする */
                bw.WorkerReportsProgress = true;

                // 途中で中止できるようにする
                bw.WorkerSupportsCancellation = true;

                // イベントハンドラを追加
                bw.DoWork += new DoWorkEventHandler(DoWork);
                bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

                /* DoWorkで取得できるパラメータを指定して、処理を開始する
                   パラメータが必要なければ省略できる */
                bw.RunWorkerAsync(sfCount);
            }
            else
            {
                /* キャンセルする */
                bw.CancelAsync();
            }

        }


        // 別スレッドで回す重たい処理（つまりはメインのループ）
        private void DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            i = 0;

             // 比較元フォルダのファイル名を配列に入れる
            string[] filenameArray = Directory.GetFiles(sPath, "*.mp3", SearchOption.TopDirectoryOnly);

            // その配列をforeachに投げ込んで処理
            foreach (string filename in filenameArray)
            {
                // キャンセルされたかどうかチェック
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // パスを除いたファイル名と、ID3タグの作曲者（ここにオリジナルのmp2ファイル名を入れている）を取得
                    // 【20190819】taglibを利用して、作曲者名を数字じゃなく要素として取得してみよう計画！
                    string oneFileName = Path.GetFileName(filename);
                    FolderItem item = f.ParseName(oneFileName);
                    string id3number = f.GetDetailsOf(item, fpValue);
                    processFile = oneFileName + Environment.NewLine;

                    // 両者のファイル名部分を比較して、違っていたら
                    if (Path.GetFileNameWithoutExtension(filename) != Path.GetFileNameWithoutExtension(id3number))
                    {
                        // テキストボックスに表示するため、直接コントロールをいじらないでListに入れておく
                        resultStrList.Add(oneFileName + " (Tag : " + id3number + ")" + Environment.NewLine);
                    }

                    i++;
                    // BackgroundWorkerに現状を送る
                    worker.ReportProgress(i);
                }
            }

            filenameArray.Initialize();
        }


        // BackgroundWorker1のProgressChangedイベントハンドラ
        // コントロールの操作は必ずここで行い、DoWorkでは絶対にしない
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // ProgressBar1の値を変更する
            MainProgressBar.Value = e.ProgressPercentage;
            ResultTextBlock.AppendText(processFile);
            //label2.Text = (e.ProgressPercentage.ToString() + "%");
        }


        // BackgroundWorker1のRunWorkerCompletedイベントハンドラ
        // 処理が終わったときに呼び出される
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((e.Cancelled == true))
            {
                textBox1.AppendText("キャンセルされました");
            }

            else if (!(e.Error == null))
            {
                textBox1.AppendText("エラー:" + e.Error.Message);
            }

            else
            {
                textBox1.AppendText("完了しました" + Environment.NewLine);

                if (resultStrList.Count > 0) // ID3Tagが違うファイルがあったら
                {
                    textBox1.AppendText("以下のファイルのID3タグが、拡張子を除くファイル名と違います" + Environment.NewLine);
                    foreach (var item in resultStrList)
                    {
                        textBox1.AppendText(item);
                    }
                }
                else
                {
                    textBox1.AppendText("問題はありません" + Environment.NewLine);
                }
            }

            // 各種初期化
            progressBar1.Visible = false;
            button2.Text = "実行";
            numericUpDown1.Enabled = true;
        }
    }





}
}
