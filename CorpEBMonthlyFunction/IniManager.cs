using System.Runtime.InteropServices;
using System.Text;

public class IniManager
{
    private string filePath;
    private StringBuilder lpReturnedString;
    private int bufferSize;

    [DllImport("kernel32")]
    private static extern long WritePrivateProfileString(string section, string key, string lpString, string lpFileName);

    [DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string section, string key, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

    public IniManager(string iniPath)
    {
        filePath = iniPath;
        bufferSize = 4096;
        lpReturnedString = new StringBuilder(bufferSize);
    }

    /// <summary>
    /// 讀取ini檔
    /// </summary>
    /// <param name="section">讀取的Section</param>
    /// <param name="key">該Section下讀取的Key值</param>
    /// <param name="defaultValue">預設值</param>
    /// <returns></returns>
    public string ReadIniFile(string section, string key, string defaultValue)
    {
        lpReturnedString.Clear();
        GetPrivateProfileString(section, key, defaultValue, lpReturnedString, bufferSize, filePath);
        return lpReturnedString.ToString();
    }

    /// <summary>
    /// 寫入ini檔
    /// </summary>
    /// <param name="section">寫入的Section</param>
    /// <param name="key">該Section下寫入的Key(若null為刪該Section)</param>
    /// <param name="value">寫入的值(若null為刪該Key)</param>
    public void WriteIniFile(string section, string key, string value)
    {
        WritePrivateProfileString(section, key, value, filePath);
    }
}
