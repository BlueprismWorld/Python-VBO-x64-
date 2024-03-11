using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Python.Runtime;

namespace PythonInterop
{
    public class PythonInteropWrapper
    {
        private readonly string _PythonInstallationRootLocation;
        private int _PythonVersion=0;
        private IntPtr pythonDll = IntPtr.Zero;

        public PythonInteropWrapper(string PythonInstallationRootLocation, string pythonVersion=null)
        {
            this._PythonInstallationRootLocation = PythonInstallationRootLocation;

            if (pythonVersion != null)
            {
                string[] ver = pythonVersion.Split('.');
                if (ver.Length >= 2)
                    this._PythonVersion = int.Parse(ver[0] + ver[1]);
                else throw new Exception("Invalid python version provided :" + pythonVersion);
            }

            Init();
        }
        private void Init()
        {
            SetGlobalPath();
            LoadLibrarys();
            PythonEngine.Initialize();
            if (!PythonEngine.IsInitialized) throw new Exception("Python not initialized");
        }

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool FreeLibrary(IntPtr lpFileName);

        public void ImportModule(string module)
        {
                using (Py.GIL())
                {
                    Py.Import(module);
                }
        }

        public void InitPip2()
        {
            Process process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    FileName = "cmd.exe",
                    WorkingDirectory = _PythonInstallationRootLocation,
                    WindowStyle = ProcessWindowStyle.Normal,
                    Arguments = "/c python.exe get-pip.py"
                }
            };
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.Start();
            process.WaitForExit(300000);
        }
        public bool InitPip()
        {
            bool isInit = false;
            using (Py.GIL())
            {
                try
                {
                    var po = Py.Import("pip");
                    isInit = true;
                }
                catch (Exception)
                {
                    InitPip2();
                    try 
                    {
                        var d = Py.Import("pip");
                        isInit = true;
                    }
                    catch (Exception)
                    {
                        throw new Exception("Unable initiliase pip please check in path " + _PythonInstallationRootLocation);
                    }
                }
            }
            return isInit;
        }

        public void RunFile(string filepath, bool isFile=false)
        {

            using (Py.GIL())
            {
                using (var scope = Py.CreateScope())
                {
                    string code;
                    if (isFile)
                    {
                        setPathGetModuleName(filepath);
                        code = File.ReadAllText(filepath);
                    }else code = filepath;
                    PythonEngine.Compile(code);
                    scope.Exec(code);
                }
            }
        }
        public void RefreshImports()
        {
            using (Py.GIL())
            {
                dynamic site = Py.Import("site");
                dynamic importlib = Py.Import("importlib");
                importlib.reload(site);
            }
        }
        private string setPathGetModuleName(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new Exception("Path should not be null");
            using (Py.GIL())
            {
                if (path.Contains(".py") && path.EndsWith(".py"))
                {
                    dynamic os = Py.Import("os");
                    dynamic sys = Py.Import("sys");
                    sys.path.append(os.path.dirname(os.path.expanduser(path)));
                    return Path.GetFileNameWithoutExtension(path);
                }
                else
                    throw new Exception("Invalid path or the path doesn't contains the file name(.py)");
            }
        }

        public T getTextFromFile<T>(string filepath, string function_name, DataTable param = null)
        {
            if (string.IsNullOrEmpty(filepath) && string.IsNullOrEmpty(function_name))
                throw new Exception("File path or Function Name should not be null");
            T obj;
                using (Py.GIL())
                {
                    PythonEngine.Compile("", filepath, RunFlagType.File);
                    var df = Py.Import(setPathGetModuleName(filepath));
                    if (param != null && param.Rows.Count >= 1)
                    {
                        var pyobj = df.InvokeMethod(function_name, ParseParams(param));
                        //return (T)pyobj.AsManagedObject(typeof(T));
                        obj = (T)pyobj.AsManagedObject(typeof(T));
                    }
                    else
                    {
                        var pyobj = df.InvokeMethod(function_name);
                        //return (T)pyobj.AsManagedObject(typeof(T));
                        obj = (T)pyobj.AsManagedObject(typeof(T));
                    }
                }
            return obj;
        }

        public T getTextFromScript<T>(string code, string function_name, DataTable dt = null)
        {
            if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(function_name))
                throw new Exception("Python script or Function Name should not be null");
            T retObj;
                using (Py.GIL())
                {
                    PythonEngine.Compile(code);
                    using (var scope = Py.CreateScope())
                    {
                        scope.Exec(code);
                        using (var scopeB = Py.CreateScope())
                        {
                            scopeB.Import(scope, "ds");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt.Rows)
                                {
                                    int count = 0;
                                    foreach (DataColumn dc in dt.Columns)
                                    {
                                        scopeB.Set("param" + count, dr[dc]);
                                        count++;
                                    }
                                    break;
                                }
                                string fux = "";
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    if (i == dt.Columns.Count - 1)
                                        fux += "param" + i;
                                    else
                                        fux += "param" + i + ",";
                                }
                                var pyobj = scopeB.Eval("ds." + function_name + "(" + fux + ")");
                                retObj = (T)pyobj.AsManagedObject(typeof(T));
                            }
                            else
                            {
                                var pyobj = scopeB.Eval("ds." + function_name + "()");
                                retObj = (T)pyobj.AsManagedObject(typeof(T));
                            }
                        }
                    }

                }
            return retObj;
            
        }



        public void setPath(string path)
        {
            using (Py.GIL())
            {
                dynamic ps = Py.Import("sys");
                ps.path.append(path);
            }
        }

        public PyObject[] ParseParams(DataTable dt)
        {
            PyObject[] pyObjects = new PyObject[dt.Columns.Count];
            foreach (DataRow dr in dt.Rows)
            {
                int count = 0;
                foreach (DataColumn dc in dt.Columns)
                {
                    pyObjects[count] = dr[dc].ToPython();
                    count++;
                }
                break;
            }
            return pyObjects;
        }
        
        private void SetGlobalPath()
        {
            string folder_path = _PythonInstallationRootLocation;
            string path = folder_path + ";" + Path.Combine(folder_path, "Scripts") + ";" + Path.Combine(folder_path, @"Lib\site-packages") + ";" + 
            Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine);
            Environment.SetEnvironmentVariable("PATH", path, EnvironmentVariableTarget.Process);
            Environment.SetEnvironmentVariable("PYTHONHOME", folder_path, EnvironmentVariableTarget.Process);
            Environment.SetEnvironmentVariable("PYTHONPATH", Path.Combine(folder_path, @"Lib\site-packages"), EnvironmentVariableTarget.Process);
            
        }

        private void LoadLibrarys()
        {
            if (_PythonVersion == 0)
            {
                int pythonIndex = _PythonInstallationRootLocation.LastIndexOf("Python");
                if (pythonIndex != -1)
                {
                    string versionString = _PythonInstallationRootLocation.Substring(pythonIndex + "Python".Length);
                    string version = new string(versionString.Where(char.IsDigit).ToArray());
                    _PythonVersion = int.Parse(versionString);
                }
                else
                {
                    throw new Exception("Unable to extract python version from path");
                }
            }
            string pythonPath = Path.Combine(_PythonInstallationRootLocation, "python"+_PythonVersion+".dll");
            if (File.Exists(pythonPath))
            {
                IntPtr h = LoadLibrary(pythonPath);
                if (h == IntPtr.Zero)
                {
                    throw new DllNotFoundException("Unable to load library: " + pythonPath);
                }
                pythonDll = h;
            }
            else throw new Exception(string.Format("The pythonDll.dll file not found in path {0}", _PythonInstallationRootLocation));
            
        }

        ~PythonInteropWrapper()
        {
            if(pythonDll != IntPtr.Zero)
            FreeLibrary(pythonDll);
        }
    }
    
}
