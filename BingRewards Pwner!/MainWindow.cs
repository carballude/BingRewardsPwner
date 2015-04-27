using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BingRewardsPwner
{
    /// <summary>
    /// I've created this tool for fun. Yes, could take advantage of Bing Rewards but... Bing is a good search engine, give it a shot!
    /// </summary>
    public partial class MainWindow : Form
    {

        private const string BING_REWARDS_LOGIN_URL = "https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=12&ct=1430103040&rver=6.0.5286.0&wp=MBI&wreply=https:%2F%2Fwww.bing.com%2Fsecure%2FPassport.aspx%3Frequrl%3Dhttps%253a%252f%252fwww.bing.com%253a443%252frewards%252fdashboard&lc=1033&id=264960";

        /// <summary>
        /// We want to identify ourselves as a mobile device to perform mobile searches, so we will use this as our User-Agent
        /// </summary>
        private const string MOBILE_USER_AGENT = "User-Agent: Mozilla 5.0 (Linux; U; Android 2.3.7; zh-cn; MB525 Build MIUI) UC AppleWebKit 534.31 (KHTML, like Gecko) Mobile Safari 534.31";

        /// <summary>
        /// Surprisingly enough, there is no easy way to permanently change the User-Agent of the WebBrowser component. The built-in way only last for a request.
        /// UrlMkSetSessionOption allows us to set the User-Agent for the session. This shouln't affect other instances of IE, but it would be nice to test it.
        /// More info: https://msdn.microsoft.com/en-us/library/ie/ms775125%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        /// </summary>
        /// <param name="dwOption">An unsigned long integer value that contains the option to set.</param>
        /// <param name="pBuffer">A pointer to the buffer containing the new session settings.</param>
        /// <param name="dwBufferLength">An unsigned long integer value that contains the size of pBuffer.</param>
        /// <param name="dwReserved">Reserved. Must be set to 0. (I wonder why...)</param>
        /// <returns></returns>
        [DllImport("urlmon.dll", CharSet = CharSet.Ansi)]
        private static extern int UrlMkSetSessionOption(int dwOption, string pBuffer, int dwBufferLength, int dwReserved);

        /// <summary>
        /// Sets the user agent string for this process.
        /// </summary>
        const int URLMON_OPTION_USERAGENT = 0x10000001;

        private bool dashboardFound = false;
        private bool searchingMobile = false;
        private string searchTerm;
        private Timer timer;

        public MainWindow()
        {
            timer = new Timer();
            timer.Tick += Timer_DesktopSearchTick;
            InitializeComponent();
            webBrowser1.DocumentCompleted += WebBrowser1_DocumentCompleted;
        }

        private int CreditTextToSearches(string text)
        {
            var chunks = text.Split(new char[] { ' ' });
            var maximum = Convert.ToInt32(chunks[2]);
            var current = Convert.ToInt32(chunks.First());
            var numberOfSearches = (maximum - current) * 2;
            return numberOfSearches;
        }

        private void DoSearch()
        {
            timer.Interval = new Random(DateTime.Now.Millisecond).Next((int)minimumSecondsUpDown.Value * 1000, (int)maximumSecondsUpDown.Value * 1000);
            timer.Start();
        }

        private void Timer_DesktopSearchTick(object sender, EventArgs e)
        {
            timer.Stop();
            webBrowser1.Navigate("http://www.bing.com/search?q=" + searchTerm);
        }

        private void DealWithWikipedia(WebBrowserDocumentCompletedEventArgs e)
        {
            searchTerm = e.Url.ToString().Substring(29).Replace('_', ' ');
            DoSearch();
        }

        private void ChangeUserAgent()
        {
            UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, MOBILE_USER_AGENT, MOBILE_USER_AGENT.Length, 0);
        }

        private void DealWithBingResults()
        {
            if (!searchingMobile)
            {
                desktopSearchesNumericUpDown.Value--;
                if (desktopSearchesNumericUpDown.Value < 0)
                {
                    searchingMobile = true;
                    ChangeUserAgent();
                }
            }
            else
                mobileSearchesNumericUpDown.Value--;
            if (mobileSearchesNumericUpDown.Value > 0 || desktopSearchesNumericUpDown.Value > 0)
                GetNewSearchTerm();
            else
            {
                startButton.Enabled = true;
                MessageBox.Show("All searches done. You've achieved all your search points for today.", "Bing Rewards Pwner!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.ToString().StartsWith("http://www.bing.com/search?q="))
            {
                DealWithBingResults();
            }
            if (e.Url.ToString().StartsWith("http://en.wikipedia.org/wiki/") || e.Url.ToString().StartsWith("http://en.m.wikipedia.org/wiki/"))
            {
                DealWithWikipedia(e);
            }
            if (e.Url.ToString().StartsWith("https://www.bing.com/msasignin"))
                FillLoginScreen();
            if (!dashboardFound)
            {
                var mobileSearchInfo = webBrowser1.Document.GetElementById("mobsrch01");
                if (mobileSearchInfo != null)
                {
                    var mobileProgress = ElementsByClass(mobileSearchInfo, "progress");
                    if (mobileProgress != null)
                    {
                        var value = mobileProgress.First().InnerText;
                        if (value.Contains(" of "))
                            mobileSearchesNumericUpDown.Value = CreditTextToSearches(value);
                        else
                            mobileSearchesNumericUpDown.Value = 0;

                    }
                }
                var pcSearchInfo = webBrowser1.Document.GetElementById("srchSMR1216");
                if (pcSearchInfo != null)
                {
                    var pcProgress = ElementsByClass(pcSearchInfo, "progress");
                    if (pcProgress != null)
                    {
                        var value = pcProgress.First().InnerText;
                        if (value.Contains(" of "))
                            desktopSearchesNumericUpDown.Value = CreditTextToSearches(value);
                        else
                            desktopSearchesNumericUpDown.Value = 0;
                    }
                }
                dashboardFound = desktopSearchesNumericUpDown.Value != -1 && mobileSearchesNumericUpDown.Value != -1;
                if (dashboardFound)
                    GetNewSearchTerm();
            }
        }

        private void GetNewSearchTerm()
        {
            webBrowser1.Navigate("http://en.wikipedia.org/wiki/Special:Random");
        }

        private void FillLoginScreen()
        {
            var user = webBrowser1.Document.GetElementById("i0116");
            if (user != null)
                user.SetAttribute("value", textBox1.Text);
            var password = webBrowser1.Document.GetElementById("i0118");
            if (password != null)
                password.SetAttribute("value", textBox2.Text);
            var button = webBrowser1.Document.GetElementById("idSIButton9");
            if (button != null && password != null && user != null)
                button.InvokeMember("Click");
        }

        /// <summary>
        /// Retrieves all elements with an specific class name
        /// </summary>
        /// <param name="element"></param>
        /// <param name="className"></param>
        /// <returns></returns>
        private IEnumerable<HtmlElement> ElementsByClass(HtmlElement element, string className)
        {
            foreach (HtmlElement e in element.Children)
                if (e.GetAttribute("className") == className)
                    yield return e;
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(BING_REWARDS_LOGIN_URL);
            startButton.Enabled = false;
        }
    }
}
