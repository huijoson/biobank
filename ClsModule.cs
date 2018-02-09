using System;

public class ShareModule
{
    public ShareModule()
	{
          /*密碼驗證: (1)長度大於8 (2)大小寫 (3)英數字*/
         Boolean gfunCheckPwd(string Pwd)
        {
            Boolean Check = false;
            char  [] Character;
            int UpperSum = 0;
            int LowerSum = 0;
            int NumSum = 0;

            if (Pwd.Length >= 8)
            {
                Character = new char[Pwd.Length];

                for (int i = 0; i < Pwd.Length; i++)
                {
                    Character[i] = Convert.ToChar(Pwd.Substring(i, 1));
                    if (Char.IsUpper(Character[i]) == true)
                        UpperSum += 1;
                    if (Char.IsLower(Character[i]) == true)
                        LowerSum += 1;
                    if (Char.IsNumber(Character[i]) == true)
                        NumSum += 1;
                }

                if (UpperSum > 0 && LowerSum > 0 && NumSum > 0)
                    Check = true;
            }

            return Check;
        }
	}
}
