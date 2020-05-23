using Microsoft.Win32;
using System;
using System.IO;
using System.Text;
namespace AddTerminalToMenu
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var Terminal_reg_path = "Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages";
            //var Terminal_family_reg_path = "Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Families";
            var Terminal_reg_name = "Microsoft.WindowsTerminal";
            var Terminal_name = "WindowsTerminal.exe";
            var Terminal_full_name = "";
            var menu_path = "\\Directory\\Background\\shell";
            var shell_name = "WindowsTerminal";
            //var family_name = "";
            var Shell_name = @"C:\Program Files (x86)\TerminalShell";
            var store_path = "ms-windows-store://pdp/?productid=9n0dx20hk701";
            var reg = Registry.CurrentUser.OpenSubKey(Terminal_reg_path, false);
            var names = reg.GetSubKeyNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].Contains(Terminal_reg_name))
                {
                    var Terminal_reg = reg.OpenSubKey(names[i], false);
                    var path = Terminal_reg.GetValue("PackageRootFolder");
                    Terminal_full_name = path.ToString() + "\\" + Terminal_name;
                    break;
                }
                if (i != names.Length - 1) continue;
                Console.WriteLine("未安装WindowsTerminal！是否安装[Y/N]");
                var input = Console.ReadLine();
                if (input.ToUpper() == "Y")
                {
                    System.Diagnostics.Process.Start(store_path);
                }
                return;
            }
            /*reg = Registry.CurrentUser.OpenSubKey(Terminal_family_reg_path, false);
            names = reg.GetSubKeyNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].Contains(Terminal_reg_name))
                {
                    family_name = names[i];
                    break;
                }
            }*/
            var menu_reg = Registry.ClassesRoot.OpenSubKey(menu_path, true);
            names = menu_reg.GetSubKeyNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].Contains(shell_name))
                {
                    Console.WriteLine("已安装右键菜单！是否卸载[Y/N]");
                    var _in = Console.ReadLine();
                    if (_in.ToUpper() == "Y")
                    {
                        menu_reg.DeleteSubKeyTree(shell_name);
                        if (Directory.Exists(Shell_name))
                            Directory.Delete(Shell_name, true);
                        Console.WriteLine("成功卸载！");
                        Console.Read();
                    }
                    return;
                }
                if (i != names.Length - 1) continue;
                Console.WriteLine("已安装WindowsTerminal是否安装右键菜单？[Y/N]");
                var input = Console.ReadLine();
                if (input.ToUpper() == "Y")
                {
                    menu_reg.CreateSubKey(shell_name);
                    menu_reg = menu_reg.OpenSubKey(shell_name, true);
                    menu_reg.SetValue("Icon", Terminal_full_name);
                    menu_reg.CreateSubKey("command");
                    menu_reg = menu_reg.OpenSubKey("command", true);
                    /*生成vbs文件*/
                    string vbs = "if right(WScript.Arguments(0),1)=\"\\\"then\n" +
                        "p=\"wt.exe\"+\" -d \"+left(WScript.Arguments(0),2)\n" +
                        "else\n" +
                        "p = \"wt.exe\"+\" -d \"\"\"+WScript.Arguments(0)+\"\"\"\"\n" +
                        "end if\n" +
                        "Set os = CreateObject(\"wscript.shell\")\n" +
                        "os.run p";
                    if (!Directory.Exists(Shell_name))
                        Directory.CreateDirectory(Shell_name);
                    File.WriteAllText(Shell_name + "\\Shell.vbs", vbs, Encoding.ASCII);
                    menu_reg.SetValue("", "wscript.exe " + "\"" + Shell_name + "\\Shell.vbs\"" + " \"%V\"");
                    Console.WriteLine("安装右键菜单成功!");
                }
                Console.Read();
            }

        }
    }
}
