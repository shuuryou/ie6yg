﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace yomigaeri_backend {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("yomigaeri_backend.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;!doctype html&gt;
        ///&lt;html lang=&quot;en&quot;&gt;
        ///&lt;head&gt;
        ///    &lt;meta charset=&quot;utf-8&quot;&gt;
        ///    &lt;meta name=&quot;viewport&quot; content=&quot;width=device-width, initial-scale=1&quot;&gt;
        ///
        ///    &lt;title&gt;%TITLE_TEX%&lt;/title&gt;
        ///
        ///    &lt;style type=&quot;text/css&quot;&gt;
        ///        body {
        ///            background-color: white;
        ///            color: black;
        ///            font-family: Verdana;
        ///            font-size: 8pt;
        ///            line-height: 11pt;
        ///        }
        ///
        ///        a:link {
        ///            color: red;
        ///        }
        ///
        ///        a:visited {
        ///            color: #4e4e4e;
        ///       [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string BACKENDERROR_HTML {
            get {
                return ResourceManager.GetString("BACKENDERROR_HTML", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;!doctype html&gt;
        ///&lt;html lang=&quot;en&quot;&gt;
        ///&lt;head&gt;
        ///    &lt;meta charset=&quot;utf-8&quot;&gt;
        ///    &lt;meta name=&quot;viewport&quot; content=&quot;width=device-width, initial-scale=1&quot;&gt;
        ///
        ///    &lt;title&gt;The page cannot be displayed&lt;/title&gt;
        ///
        ///    &lt;style type=&quot;text/css&quot;&gt;
        ///        body {
        ///            background-color: white;
        ///            color: black;
        ///            font-family: Verdana;
        ///            font-size: 8pt;
        ///            line-height: 11pt;
        ///        }
        ///
        ///        a:link {
        ///            color: red;
        ///        }
        ///
        ///        a:visited {
        ///            color: [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string BROWSERERROR_HTML {
            get {
                return ResourceManager.GetString("BROWSERERROR_HTML", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to There was a problem trying to communicate with the Internet Explorer frontend and the communication channel was closed. Your computer may be experiencing network errors. Any information you were working on has been lost..
        /// </summary>
        internal static string E_BackendVirtualChannel_Text {
            get {
                return ResourceManager.GetString("E_BackendVirtualChannel_Text", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Communication with frontend was lost.
        /// </summary>
        internal static string E_BackendVirtualChannel_Title {
            get {
                return ResourceManager.GetString("E_BackendVirtualChannel_Title", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not initialize the AdBlock engine:
        ///{0}.
        /// </summary>
        internal static string E_InitErrorAdBlock {
            get {
                return ResourceManager.GetString("E_InitErrorAdBlock", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not initialize the Chromium engine:
        ///{0}.
        /// </summary>
        internal static string E_InitErrorChromiumEmbeddedFramework {
            get {
                return ResourceManager.GetString("E_InitErrorChromiumEmbeddedFramework", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not open the settings file &quot;{0}&quot;:
        ///{1}.
        /// </summary>
        internal static string E_InitErrorCouldNotLoadSettings {
            get {
                return ResourceManager.GetString("E_InitErrorCouldNotLoadSettings", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not open the log file &quot;{0}&quot;:
        ///{1}.
        /// </summary>
        internal static string E_InitErrorCouldNotOpenLog {
            get {
                return ResourceManager.GetString("E_InitErrorCouldNotOpenLog", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not open the virtual channel to the frontend:
        ///{0}.
        /// </summary>
        internal static string E_InitErrorCouldNotOpenRDPVC {
            get {
                return ResourceManager.GetString("E_InitErrorCouldNotOpenRDPVC", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Sorry, the backend is meant to be used via Remote Desktop only..
        /// </summary>
        internal static string E_InitErrorNotTerminalSession {
            get {
                return ResourceManager.GetString("E_InitErrorNotTerminalSession", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Initialization Error.
        /// </summary>
        internal static string E_InitErrorTitle {
            get {
                return ResourceManager.GetString("E_InitErrorTitle", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Opening page {0}....
        /// </summary>
        internal static string StatusConnecting {
            get {
                return ResourceManager.GetString("StatusConnecting", resourceCulture);
            }
        }
    }
}