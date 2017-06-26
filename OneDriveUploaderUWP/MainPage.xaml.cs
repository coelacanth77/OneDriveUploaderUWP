using Microsoft.Graph;
using OneDriveUploaderUWP.Models.Helpers;
using System;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media.Imaging;

// 空白ページの項目テンプレートについては、https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x411 を参照してください

namespace OneDriveUploaderUWP
{
    /// <summary>
    /// それ自体で使用できる空白ページまたはフレーム内に移動できる空白ページ。
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private GraphServiceClient graphClient { get; set; }

        private DriveItem currentFolder { get; set; }

        private IRandomAccessStream stream { get; set; }

        public MainPage()
        {
            this.InitializeComponent();

            this.Loaded += MainPage_Loaded;
        }

        private void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            this.selectImageButton.Visibility = Visibility.Collapsed;
            this.uploadImageButton.Visibility = Visibility.Collapsed;
        }

        private async void connectOneDriveButton_Click(object sender, RoutedEventArgs e)
        {
            // OneDriveを操作するためのクラスを取得する
            // OneDriveはGraph APIを用いるのでgraphという名前が頻出します。
            this.graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (this.graphClient == null)
            {
                messageText.Text = "OneDriveとの接続に失敗しました。AppIDnなどが間違っていないか確認ください。";
            }

            // OneDriveのrootフォルダーの情報を取得する
            // 認証が済んでいない場合はログイン画面を表示する
            this.currentFolder = await this.graphClient.Drive.Root.Request().Expand("").GetAsync();

            if (this.currentFolder != null)
            {
                // 接続に成功したら画像を選択させるためのボタンを表示する
                this.connectOneDriveButton.Visibility = Visibility.Collapsed;
                this.selectImageButton.Visibility = Visibility.Visible;
            
                this.messageText.Text = "OneDriveに接続されました。アップロードする画像を選択してください。";
            }
        }

        private async void selectImageButton_Click(object sender, RoutedEventArgs e)
        {
            // 画像を選択する処理
            FileOpenPicker openPicker = new FileOpenPicker();

            openPicker.ViewMode = PickerViewMode.Thumbnail;
            openPicker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;

            openPicker.FileTypeFilter.Add(".jpg");
            openPicker.FileTypeFilter.Add(".jpeg");

            StorageFile file = await openPicker.PickSingleFileAsync();

            if (file != null)
            {
                this.stream = await file.OpenAsync(Windows.Storage.FileAccessMode.ReadWrite);

                BitmapImage imageSource = new BitmapImage();
                imageSource.SetSource(this.stream);

                selectedImage.Source = imageSource;

                // 画像が選択されたアップロードボタンを有効にする
                this.uploadImageButton.Visibility = Visibility.Visible;
            }
        }

        private void uploadImageButton_Click(object sender, RoutedEventArgs e)
        {
            var targetFolder = this.currentFolder;

            var fileName = "sample.jpg";

            // パスを生成する
            string folderPath = targetFolder.ParentReference == null
            ? ""
            : targetFolder.ParentReference.Path.Remove(0, 12) + "/" + Uri.EscapeUriString(targetFolder.Name);
            var uploadPath = folderPath + "/" + Uri.EscapeUriString(System.IO.Path.GetFileName(fileName));

            // シーク位置を最初に戻しておく
            this.stream.Seek(0);

            // OneDriveにアップロードする
            this.graphClient.Drive.Root.ItemWithPath(uploadPath).Content.Request().PutAsync<DriveItem>(this.stream.AsStream());

            this.messageText.Text = "sample.jpgという名前で画像が保存されました";
        }
    }
}
