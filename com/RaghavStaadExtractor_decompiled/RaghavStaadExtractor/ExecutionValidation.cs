using System;
using System.IO;
using System.Management;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace RaghavStaadExtractor;

[ComVisible(false)]
internal class ExecutionValidation
{
	internal class LicenseCheckResult
	{
		internal LicenseStatus Status { get; set; }

		internal string ExpiryDate { get; set; }
	}

	internal enum LicenseStatus
	{
		Valid,
		Expired,
		Unauthorized,
		InvalidFormat
	}

	private readonly string aesKey;

	private readonly string aesIV;

	private readonly string[] fakeErrors;

	public ExecutionValidation()
	{
		aesKey = "AshuVedantRaghav";
		aesIV = "AshuVedantRaghav";
		fakeErrors = new string[10] { "Runtime Error 0x80070005: Access is denied.", "Unhandled Exception: System.IO.FileLoadException", "Fatal Error: Missing dependency 'mscorlib.dll'.", "Memory read violation at address 0x00000014.", "System.NullReferenceException at MainModule()", "Unexpected token in config file at line 42.", "Access violation while reading registry key.", "System.BadImageFormatException: Invalid executable format.", "Fatal: CLR runtime initialization failed.", "Exception code 0xc0000005: Access violation." };
	}

	internal bool IsLicenceValid()
	{
		bool result2;
		try
		{
			DateTime now = DateTime.Now;
			DateTime result = DateTime.MinValue;
			string path = "C:\\PlateNova\\PlateNova.tst";
			if (!File.Exists(path))
			{
				goto IL_0073;
			}
			string encryptedBase = File.ReadAllText(path);
			string s = AES_Decrypt(encryptedBase);
			if (!DateTime.TryParse(s, out result))
			{
				ShowFakeError();
				result2 = false;
			}
			else
			{
				if (DateTime.Compare(now, result) >= 0)
				{
					goto IL_0073;
				}
				ShowFakeError();
				result2 = false;
			}
			goto end_IL_0001;
			IL_0073:
			string text = GetBiosSerial().ToUpper();
			bool flag;
			if (string.IsNullOrWhiteSpace(text))
			{
				ShowFakeError();
				result2 = false;
			}
			else
			{
				flag = false;
				DateTime t = new DateTime(2025, 12, 31);
				if (DateTime.Compare(DateTime.Today, t) <= 0)
				{
					flag = true;
					goto IL_0125;
				}
				string text2 = "C:\\PlateNova\\ExecutionFile.lic";
				if (!File.Exists(text2))
				{
					ShowFakeError();
					result2 = false;
				}
				else
				{
					LicenseCheckResult licenseCheckResult = ParseLicenseFile(text2, text);
					if (licenseCheckResult.Status == LicenseStatus.Valid)
					{
						flag = true;
						goto IL_0125;
					}
					ShowFakeError();
					result2 = false;
				}
			}
			goto end_IL_0001;
			IL_0125:
			if (flag)
			{
				try
				{
					string plainText = now.ToString("yyyy-MM-dd HH:mm:ss");
					string text3 = AES_Encrypt(plainText);
					if (!string.IsNullOrEmpty(text3))
					{
						File.WriteAllText(path, text3);
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				result2 = true;
			}
			else
			{
				result2 = false;
			}
			end_IL_0001:;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ShowFakeError();
			result2 = false;
			ProjectData.ClearProjectError();
		}
		return result2;
	}

	private LicenseCheckResult ParseLicenseFile(string licPath, string biosSerial)
	{
		LicenseCheckResult result;
		try
		{
			string[] array = File.ReadAllLines(licPath);
			string[] array2 = array;
			int num = 0;
			while (true)
			{
				if (num < array2.Length)
				{
					string encryptedBase = array2[num];
					string text = AES_Decrypt(encryptedBase).Trim();
					if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith("#") && !text.StartsWith("//") && text.Contains("="))
					{
						string[] array3 = text.Split('=');
						if (array3.Length == 2)
						{
							string right = array3[0].Trim().ToUpper();
							string text2 = array3[1].Trim().ToLower();
							if (Operators.CompareString(biosSerial, right, TextCompare: false) == 0)
							{
								result = ((!((Operators.CompareString(text2, "forever", TextCompare: false) == 0) | (Operators.CompareString(text2, "permanent", TextCompare: false) == 0))) ? ((!DateTime.TryParse(text2, out var result2)) ? new LicenseCheckResult
								{
									Status = LicenseStatus.InvalidFormat,
									ExpiryDate = text2
								} : ((DateTime.Compare(DateTime.Today, result2) > 0) ? new LicenseCheckResult
								{
									Status = LicenseStatus.Expired,
									ExpiryDate = result2.ToString("yyyy-MM-dd")
								} : new LicenseCheckResult
								{
									Status = LicenseStatus.Valid,
									ExpiryDate = result2.ToString("yyyy-MM-dd")
								})) : new LicenseCheckResult
								{
									Status = LicenseStatus.Valid,
									ExpiryDate = "Forever"
								});
								break;
							}
						}
					}
					num = checked(num + 1);
					continue;
				}
				result = new LicenseCheckResult
				{
					Status = LicenseStatus.Unauthorized,
					ExpiryDate = ""
				};
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = new LicenseCheckResult
			{
				Status = LicenseStatus.InvalidFormat,
				ExpiryDate = ex2.Message
			};
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string AES_Decrypt(string encryptedBase64)
	{
		string result;
		try
		{
			byte[] bytes = Encoding.UTF8.GetBytes(aesKey);
			byte[] bytes2 = Encoding.UTF8.GetBytes(aesIV);
			byte[] buffer = Convert.FromBase64String(encryptedBase64);
			using Aes aes = Aes.Create();
			aes.Key = bytes;
			aes.IV = bytes2;
			aes.Padding = PaddingMode.PKCS7;
			using ICryptoTransform transform = aes.CreateDecryptor();
			using MemoryStream stream = new MemoryStream(buffer);
			using CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Read);
			using StreamReader streamReader = new StreamReader(stream2);
			result = streamReader.ReadToEnd();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string AES_Encrypt(string plainText)
	{
		string result;
		try
		{
			byte[] bytes = Encoding.UTF8.GetBytes(aesKey);
			byte[] bytes2 = Encoding.UTF8.GetBytes(aesIV);
			using Aes aes = Aes.Create();
			aes.Key = bytes;
			aes.IV = bytes2;
			aes.Padding = PaddingMode.PKCS7;
			using ICryptoTransform transform = aes.CreateEncryptor();
			using MemoryStream memoryStream = new MemoryStream();
			using (CryptoStream stream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Write))
			{
				using StreamWriter streamWriter = new StreamWriter(stream);
				streamWriter.Write(plainText);
			}
			result = Convert.ToBase64String(memoryStream.ToArray());
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string GetBiosSerial()
	{
		try
		{
			ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS");
			foreach (ManagementObject item in managementObjectSearcher.Get())
			{
				string text = item["SerialNumber"].ToString().Trim();
				if (!string.IsNullOrWhiteSpace(text))
				{
					return text;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return "";
	}

	private void ShowFakeError()
	{
		try
		{
			Random random = new Random();
			string prompt = fakeErrors[random.Next(fakeErrors.Length)];
			Interaction.MsgBox(prompt, MsgBoxStyle.Critical, "Fatal Error");
			try
			{
				object obj = null;
				try
				{
					obj = RuntimeHelpers.GetObjectValue(Marshal.GetActiveObject("Excel.Application"));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				if (obj != null)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, "ActiveWorkbook", new object[0], null, null, null));
					if (objectValue != null)
					{
						NewLateBinding.LateCall(objectValue, null, "Close", new object[1] { false }, null, null, null, IgnoreReturn: true);
					}
					Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(obj));
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
	}

	internal string GetCurrentBiosSerial()
	{
		return GetBiosSerial();
	}

	internal bool ValidateLicenseFileFormat(string licPath)
	{
		bool result;
		try
		{
			if (!File.Exists(licPath))
			{
				result = false;
			}
			else
			{
				string[] array = File.ReadAllLines(licPath);
				string[] array2 = array;
				int num = 0;
				while (true)
				{
					if (num < array2.Length)
					{
						string encryptedBase = array2[num];
						string text = AES_Decrypt(encryptedBase).Trim();
						if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith("#") && text.Contains("="))
						{
							string[] array3 = text.Split('=');
							if (array3.Length == 2 && !string.IsNullOrWhiteSpace(array3[0]) && !string.IsNullOrWhiteSpace(array3[1]))
							{
								result = true;
								break;
							}
						}
						num = checked(num + 1);
						continue;
					}
					result = false;
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
