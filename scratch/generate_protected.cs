using System;
using System.Text;

public class Program {
    public static void Main() {
        string text = "https://raw.githubusercontent.com/mdshahnawaz123/plugin-access-control/main/users.json";
        int Key = 0xBDD;
        var result = new StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            result.Append((char)(text[i] ^ Key));
        }
        Console.WriteLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(result.ToString())));
    }
}
