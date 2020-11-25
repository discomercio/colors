using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.EnterpriseServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("ComPlusWrapper_DllInscE32")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("ComPlusWrapper_DllInscE32")]
[assembly: AssemblyCopyright("Copyright ©  2019")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(true)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("0DE23D48-4B6E-4852-88D2-8AC3958EC57C")]

[assembly: ApplicationActivation(ActivationOption.Server)]
// Ao instalar, faz c/ que fique desativada a seguinte opção em Serviços de Componente:
//     Propriedades -> Segurança -> Autorização -> Aplicar verificações de acesso neste aplicativo
[assembly: ApplicationAccessControl(false)]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
[assembly: AssemblyVersion("1.01.0.0")]
[assembly: AssemblyFileVersion("1.01.0.0")]
