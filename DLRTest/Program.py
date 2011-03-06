import clr
import System
clr.LoadAssemblyFromFileWithPath("E:\\My Documents\\Visual Studio 2010\\Projects\\DynamicXml\\DynamicXml\\bin\\Debug\\DynamicXml.dll")
from DynamicXml import *
from System import *
Dn.File = XDynamicExtensions.EmptyXDynamic
Dn.File.FilePath = "D:\out\hosts.txt"
Dn.File.ReplaceRules = XDynamicExtensions.EmptyXDynamic
Dn.File.ReplaceRules.Search = "$CONSUMER_TOKEN$"
Dn.File.ReplaceRules.Replace = "asddasfsdfawdvcasdvc"
Dn.File.ReplaceRules(1).Search = "sdsdfa"
Dn.File.ReplaceRules(1).Replace = "$CONSUMER_SECRET$"
print Dn.ToString()