using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Renci.SshNet;
using System.Text.RegularExpressions;
using System.Threading;
using System.Runtime.InteropServices;



namespace SSH4Excel
{

   [ComVisible(true), Guid("5ECC6EA1-4BBA-44B6-BB83-097CD0462A90")]
   [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISSH
   {
       Boolean open(String host, String user, String pass, String enable);
        void close();
        Boolean prompt(String reg);
        void timer(Int16 sec);
        String expect(String cmd);
        String expect(String cmd, int t);
        String read();
        void write(String msg);
        String pureExpect(Regex msg);
        String expectPrompt();
   }

   [ComVisible(true), Guid("7CBEE70B-53D5-4C2D-8905-8100B8968BBB"),
    ClassInterface(ClassInterfaceType.None)]
   [ComDefaultInterface(typeof(ISSH))]
    public class SSH : ISSH
   {
        ConnectionInfo connInfo;
        SshClient session;
        Regex reg_prompt = new Regex(@"[>#]");
        ShellStream stream;
        TimeSpan waitTime = new TimeSpan(0, 0, 2);


        public SSH() { }

        public Boolean open(String host, String user, String pass, String enable)
        {
            try
            {
                connInfo = new ConnectionInfo(host, 22, user,
                    new AuthenticationMethod[]{
                    new PasswordAuthenticationMethod(user, pass)
                    }
                );
                session = new SshClient(connInfo);
                session.Connect();
                stream = session.CreateShellStream("dumb", 0, 0, 0, 0, 8192);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void close()
        {
            session.Disconnect();
        }


        public Boolean prompt(String reg)
        {
            reg_prompt = new Regex(reg);
            return true;
        }

        public void timer(Int16 sec)
        {
            waitTime = new TimeSpan(0, 0, sec);
        }


        public String expect(String cmd, int t)
        {
            String retval = "";
            try 
            {
                stream.Write(cmd + "\r");
                stream.Flush();
                Thread.Sleep(t);
                retval = stream.Expect(reg_prompt, waitTime);
                stream.Flush();
                while (stream.DataAvailable == true)
                {
                    retval += stream.Read();
                }
                return retval;
            }
            catch (System.Exception e)
            {
                return retval + "\n" + "err: " + e.ToString();
            }
        }

        public String expect(String cmd)
        {
            return this.expect(cmd, 0);
        }

        public String pureExpect(Regex msg)
        {
            String retval = "";

            stream.Flush();
            stream.Expect(msg, waitTime);
            stream.Flush();
            while (stream.DataAvailable == true)
            {
                retval += stream.Read();
            }
            return retval;
        }

        public String expectPrompt()
        {
            return this.pureExpect(reg_prompt);
        }




        public String read()
        {
            return stream.Read();
        }

        public void write(String msg)
        {
            stream.Write(msg);
        }



        public void test()
        {
//            String msg;
            Console.WriteLine(expect("terminal length 0"));
            Console.WriteLine(expect("sh run"));

        }

    }
}
