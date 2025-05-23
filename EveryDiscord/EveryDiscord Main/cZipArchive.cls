VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cZipArchive (c) 2017-2019 by wqweto@gmail.com
'
' A single-class pure VB6 library for zip archives management
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cZipArchive"

#Const ImplCompress = ZIP_NOCOMPRESS = 0
#Const ImplDecompress = ZIP_NOEXTRACT = 0
#Const ImplCrypto = ZIP_CRYPTO <> 0
#Const ImplUseShared = ZIP_USESHARED <> 0
#Const ImplInflate = ZIP_INFLATE <> 0
#Const ImplLogging = DEBUG_MODE

'=========================================================================
' Thunk data
'=========================================================================

' Auto-generated on 12.1.2018 17:15:52, CodeSize=6640, DataSize=984, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "UYtEJAhTi1wkEFWLbCQYVleLeEQD64tEJCSJbCQQiwD2wwN0HDvddBQPthNDD7bIM9HB6AgzBJf2wwN16IlcJByL1cdEJBgAAAAAK9OD4vyNDBqDwgPB6gI7yxvJ99EjyolMJCB0eYvpjaQkAAAAAIsbi8vB6QgPtvGLyMHpCA+2yTPxi8vB6RAPttGLyMHpEIu0twAIAAAPtskz0YvIwekYM7SXAAQAAIvTweoYM9EPtsgPtsOLXCQcM8iDwwQzNJeJXCQcM7SPAAwAAItMJBiLxkGJTCQYO811lItsJBCLzTP2K8s76xvt99Uj6XQWD7YTjVsBD7bIRjPRwegIMwSXO/V16otMJCRfXl2JAVtZwhAAzMzMzMzMzMzMzMzMVot0JBCF9n41i1QkDE6LRCQIwe4ERoMCAYsKdQP/QgSJCItKBIlIBMdACAAAAADHQAwAAAAAg8AQg+4BddhewhAAzMzMzMzMzMzMzMzMzMyLVCQMhdJ0GotEJAhWi3Qk" & _
    "CCvwigwGjUABMEj/g+oBdfJewhAAzMzMzMzMzMzMzMyDfCQMAItEJASLSAyJTCQED46OAAAAU1VWi3QkFFeLfCQgi++D5QGD5wKNmwAAAACKHoXtdBWLSAiDyQKL0YPyAQ+v0cHqCDLTiBaF/3QCih4PthBGD7bLi1wkFDPRiwjB6QiLFJOLWAgz0Q+2ygNIBGnJBYQICIkQD7bTwesIQYlIBA+2SAcz0YtMJBQzHJGLTCQcSYlYCIlMJByFyX+QX15dW8IQAMzMzMzMzMzMzFaLdCQIV2oMi0Yo/9CL+FeJN+gJCQAAi04oahT/0cdAEAAAAADHQAwAAAAAiUcIi8dfXsIQAMzMzMzMzMzMzMxWi3QkCFeLPv92CItHMP/Q/3YEi0cw/9CLRzBW/9BfXsIQAMzMzMzMzMzMzMzMzMyLRCQMVot0JAxXi3wkDIXAdA1Q/3YE/zb/N+gh/f///3YcjUYM/3YY/3YU/3YQUI1GCFD/dgT/NlfoIgkAAF9e" & _
    "whAAzMzMzMzMzMzMzMzMzP90JATo5w8AAMIQAMzMzMxWi3QkCItODI1GDFeLPoXJdAw7TgR0B1BW6IIRAACLThCNRhCFyXQMO04IdAdQVuhsEQAAg34UAI1GFHQHUFboXBEAAI1GBFBW6FIRAACNRghQVuhIEQAAi0cwVv/QX17CEADMzMzMzMzMzMzMzMzMi0QkCFNVi2wkDFZXjVgMU414CFf/cAT/MFXoIQkAAItMJByL8IXJdBiF9nQUiweFwHQMUf8zUP91AOgx/P//i8ZfXl1bwhAAzMzMzMzMzMxWi3QkCIPK/4uGBIAEAA+3TMYCjQTGZjvKdAkPv8FmiRTG6xCLQASD+P90CGaJlEYIgAQAi4YEgAQAi0wkEIlMxgSLhgSABABmiVTGAo0UTouOBIAEAA+3ggiABABmiQTOi4YEgAQAD78Mxg+3wGaJggiABACD+f90DGaLhgSABABmiUTOAouOBIAEAIpEJAyIhA4AAAQAi4YEgAQAQCX/" & _
    "fwAAiYYEgAQAXsIMAMzMzIPsDFNVi2wkIFZXi3wkIDP2i38EiXwkGDm3DIAGAA+O4gAAAIuPDIAGAIvBK8YDxYP4BA+MnAAAAItUJCQ78X0KD7aEPgiABgDrCIvGK8EPtgQQjV4CiEQkFI1D/zvBfQoPtoQ+CYAGAOsJi8YrwQ+2RBABiEQkFTvZfQoPtoQ+CoAGAOsJi8YrwQ+2RBACiEQkFo1DATvBfQoPtoQ+C4AGAOsJi8YrwQ+2RBADiEQkF4tEJBRpyL2nNR7B6RBRUFfoiv7//0Y7twyABgAPjFH////rMYvOO7cMgAYAfSeNlwiABgDrCY2kJAAAAACL/4qEDwiABgCNUgGIQv9BO48MgAYAfOoptwyABgCF7Q+O9QEAAOsKjaQkAAAAAI1JAItcJCSD/QQPjFkBAACLhwAABAAz7YmHAIAEAGkDvac1HsHoEA+/lEcIgAQAg/r/D4QxAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/" & _
    "fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJCx9CQ+/FNeD+v91wIXtD47pAAAAi0QkKDlEJDB+BolEJDDrBItEJDC+BAAAADvGfnGNmwAAAACKBB4zyTPbiEQkE4XtfliQi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCSKBDiLfCQYOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkMEaL64tcJCQ78HyZi1wkJItsJCg79X09i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkKHzUi2wkKFb/txCABgD/dCQo6KAOAACF9n596ziLTCQgigM8j4sRi3EID7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFZS6EcSAAC+AQAAAItsJCiD/QR8FmkDvac1HsHoEFAPtgNQV+h1/P//6xWLjwyABgCKA4iEDwiABgD/hwyA" & _
    "BgBOQ02F9n/JiWwkKIlcJCSF7Q+PF/7//19eXVuDxAzCFADMzMzMzMzMzMzMzMzMg+wUU1WLbCQoVleLfCQoM/aLfwSJfCQgObcMgAYAD47iAAAAi48MgAYAi8ErxgPFg/gED4ycAAAAi1QkLDvxfQoPtoQ+CIAGAOsIi8YrwQ+2BBCNXgKIRCQYjUP/O8F9Cg+2hD4JgAYA6wmLxivBD7ZEEAGIRCQZO9l9Cg+2hD4KgAYA6wmLxivBD7ZEEAKIRCQajUMBO8F9Cg+2hD4LgAYA6wmLxivBD7ZEEAOIRCQbi0QkGGnIvac1HsHpEFFQV+hq+///Rju3DIAGAA+MUf///+sxi847twyABgB9J42XCIAGAOsJjaQkAAAAAIv/ioQPCIAGAI1SAYhC/0E7jwyABgB86im3DIAGADPAM8mJRCQUiUwkGIlEJByF7Q+OCQIAAOsKjaQkAAAAAI1JAItcJCyD/QQPjMABAACLhwAABAAz7YmHAIAEAGkDvac1" & _
    "HsHoEA+/lEcIgAQAg/r/D4SYAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJDR9CQ+/FNeD+v91wIXtD45MAQAAi1QkMItEJDg7wn4Gi8KJRCQ4vgQAAAA7xn53igQeM8kz24hEJBOF7X5g6weNpCQAAAAAi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCyKBDiLfCQgOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkOEaL64tcJCw78HyRi1wkLItUJDA78n05i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkMHzUi2wkFIXtfl6NRQE78H5Pi0QkKIsQi2gIi0QkHDyPD7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFVS6CAPAAAPtgOLjxCABgCJTCQYiXQk" & _
    "FIlEJBzpkAAAAItEJBhVUOsli4cQgAYAiUQkGA+2A4l0JBSJRCQc63GLTCQYi2wkFIXtfjJVUf90JDDo8AoAADPAjXX/iUQkFIX2f1KLbCQwi0wkGIXtD48D/v//X15dW4PEFMIUAItMJCiKAzyPixGLcQgPtsiLQjx3CQ+2RAgwagjrDQ+2BAhqCY0ERQEAAABQVlLodg4AAL4BAAAAi2wkMIP9BHwWaQO9pzUewegQUA+2A1BX6KT4///rFYuPDIAGAIoDiIQPCIAGAP+HDIAGAE5DTYX2f8mJbCQwiVwkLOlq////zMzMzMzMzMzMVleLfCQMaLCPBgCLB4tAKP/Qi/CF9nUFX17CBACJdwSNTgK6AIAAAIPI/+sDjUkAx0EC/////41JCGaJQfhmiUH2g+oBdemNvgiABAC5AIAAAPOrX4mWBIAEAI1CAYmWDIAGAF7CBADMzMzMzMzMzFZXi3wkDLkSAAAAvgBQ51W4kCnnVfOli0wkEIPAYF9e" & _
    "xwEAEOdViUEEx0EIoEDnVcdBDHhE51XCCADMzMzMzMxTi1wkCFZXaACAAACLO4tzCItHKP/QiQYzwDlEJCRqAw+VwMdGCACAAACDwALHRgQAAAAAUFZX6DQNAACDfCQoAP90JDD/dCQw/3QkIP90JCBTdAfoF/j//+sF6DD7//9qB2oAVlfoBQ0AAIN8JCQAdBiLThCFyXQRuAgAAAArwVBqAFZX6OYMAACLRCQciw5fiQiLRCQci04EXluJCLgBAAAAwiQAzMzMzMzMzMzMzFFTVVaLdCQUV2gAgAAAix6JXCQci0Mo/9CLbCQgiYakgQAAx4asgQAAAIAAAMeGqIEAAAAAAACF7X8Ng76cAQAAAA+OtgUAAIO+nAEAABiNvpwBAAB9KotUJBzrA41JAIXtfhkPtgJNiw9C0+AJhpgBAACNQQiJB4P4GHzjiVQkHItGGIXAdQnHRhgBAAAA66aD+AF1c4sXg/oDD4xdBQAAi4aYAQAAg8L90eiLyIkX" & _
    "wegCiYaYAQAAg+EDdR2LysdGGAoAAACD4Qcr0dPoiReJhpgBAADpXP///4P5AXUYi0YEiUYMi0YIiUYQx0YYBgAAAOk/////g/kCD4U2////iU4Y6S7///+D+AJ1bosXg/oOD4zlBAAAi46YAQAAi8GD4B/B6QUFAQEAAMdGLAAAAACJRiCLwYPgH8HpBUDHRhgDAAAAiUYki8GD4A/B6QSDwASJjpgBAACJRiiNQvKJBzPAiUZEiUZIiUZMiUZQZolGVIhGVum7/v//g/gDdW85Bw+MdQQAAItGLDtGKH0xgz8DfCyLS0CLRiyKlpgBAACA4gcPtgQBiFQwRP9GLItGLIMH/cGumAEAAAM7Rih8z4tGLDtGKA+Faf7//2oTjUZEUFboxgkAAIlGFMdGGAQAAADHRiwAAAAA6Uf+//+D+AQPheIAAACLRiSLTiADwTlGLHw/UY1GV1BW6I8JAAD/diSJRgyLRiCDwFcDxlBW6HoJAACNfhSJRhBXVui9"
Private Const STR_THUNK2 As String = _
    "BQAAxwcAAAAAx0YYBgAAAOny/f///3YUjYaYAQAAV1DoGwYAAIvIg/n/D4SeAwAAg/n+D4RiAwAAg/kQfQ+LRiyITDBX/0Ys6br9//91B7gCAAAA6w8zwIP5EQ+VwI0EhQMAAACJRjAzwIP5Eg+UwI0ExQMAAACJRjSD+RB1G4tGLIXAfhQPtkQwVolGPMdGGAUAAADpbf3//zPAx0YYBQAAAIlGPOlc/f//g/gFdVyLH4tOMDvZD4wRAwAAi4aYAQAAugEAAADT4ivZSokfI9DT6ANWNImGmAEAAIXSfhyLRiQDRiCLTiw7yH0PikY8SohEMVf/RiyF0n/ki1wkGMdGGAQAAADp+/z//4P4Bg+FnAAAAP92DI1eDFeNhpgBAABQ6BgFAACD+P8PhJ0CAACD+P4PhF0CAAA9AAEAAH0QUFboCAQAAItcJBjptvz//3VAiwPHRhgBAAAAO0YEdA1TVuhYBAAAxwMAAAAAi0YQjX4Qi1wkGDtGCA+Ehvz/" & _
    "/1dW6DgEAADHBwAAAADpdPz//4tcJBg9HgEAAA+NZfz//8dGGAcAAACJRhzpVvz//4P4B3VUi0YcLQEBAACNDECLQzSNHIgPv0sCOQ8PjPsBAACLlpgBAAC4AQAAANPgSCPCA0MEiUY4D79DAikHiksCi1wkGNPqiZaYAQAAx0YYCAAAAOn9+///g/gIdTr/dhCNhpgBAABXUOghBAAAg/j/D4SmAQAAg/j+D4RqAQAAg/geD41hAQAAx0YYCQAAAIlGHOm++///g/gJD4WPAAAAi0YcixeNDECLQziNBIgPv0gCiUQkIIlMJBA70Q+MWgEAAIuGmAEAALsBAAAAK1QkENPji0wkIEsj2ANZBIkXikkC0+iDfjgAiYaYAQAAx0YYBgAAAHQsjaQkAAAAAIuGoIEAAP9OOCvDJf9/AAAPtoQwoAEAAFBW6IECAACDfjgAddv/TjiLXCQY6Sb7//+D+Ap1MYsXg/oQD4zdAAAAi46YAQAAD7fBiUZAjULw" & _
    "wekQiQeJjpgBAADHRhgLAAAA6fD6//+D+At1P4sHg/gQD4ynAAAAi46YAQAAg8DwD7fRiQeB8v//AACLRkDB6RCJjpgBAAA7wnVQ99gbwIPgC0CJRhjprPr//4P4DA+Fo/r//4M/CHxlD7aGmAEAAFBW6NkBAACDB/jBrpgBAAAIg0ZA/w+FfPr//8dGGAEAAADpcPr//4tcJBj/tqSBAACLQzD/0ItEJCTHhqSBAAAAAAAAX17HAAAAAACLRCQgXVvHAAAAAAAzwFnCFACLTCQkuAEAAACLlqSBAABfiRGLlqiBAACLTCQkXl1biRFZwhQAzMzMzMzMzMzMzMzMzIHsIAEAAFNWi7QkLAEAAFdosIEAAItGKP/Qi9iNfCQMuAgICAi5JAAAAGggAQAAiTPzq7gJCQkJjbwkoAAAALkcAAAA86uNRCQQx4QkEAEAAAcHBwdQU8eEJBwBAAAHBwcHx4QkIAEAAAcHBwfHhCQkAQAABwcHB8eEJCgBAAAH" & _
    "BwcHx4QkLAEAAAcHBwfHhCQwAQAACAgICMeEJDQBAAAICAgI6McEAACJQwSNRCQMaiBQU8dEJBgFBQUFx0QkHAUFBQXHRCQgBQUFBcdEJCQFBQUFx0QkKAUFBQXHRCQsBQUFBcdEJDAFBQUFx0QkNAUFBQXodwQAAF+JQwiLw17HQxgAAAAAx0MUAAAAAMdDEAAAAADHQwwAAAAAx4OYAQAAAAAAAMeDnAEAAAAAAADHg6CBAAAAAAAAW4HEIAEAAMIEAMzMzMzMzMzMzMzMzFOLXCQMVot0JAyLhqCBAACLDoicMKABAACLhqCBAABAJf9/AACJhqCBAACLhqiBAAA7hqyBAAB8GgPAUP+2pIEAAImGrIEAAItBLP/QiYakgQAAi4aogQAAi46kgQAAiBwB/4aogQAAXlvCCADMzMxTi1wkCFWLbCQQiwOJRCQMhe11CF2DyP9bwggAV4t9AIX/dEtWM/Y5N3wj6wONSQCLRwSNBPCDwASDOAB0B1BT" & _
    "6Lv///9GOzd+5otEJBT/dwSLQDD/0ItEJBTHRwQAAAAAV4tAMP/Qx0UAAAAAAF5fXTPAW8IIAMzMzMzMzMzMzItEJAxTi1wkDFaLCItABIszV4t8JBCLFyPKjQTID7YIO85/Jg+2CNPqK/Fmg3gC/3Usi0AEhcB0GosIi0AEI8qNBMgPtgg7zn7aX16DyP9bwgwAX164/v///1vCDACJF4kzD79AAl9eW8IMAMzMzMyD7AiLRCQMVYsoi0AIiUQkEItEJBiJbCQEhcAPji0BAABTVlc9BAEAAH4HvgIBAADrDIvwPQIBAAB+A41w/SvGux0AAACJRCQkg8//i0U0i+iJRCQUjQQ7mSvC0fiNDEA7dI0EfQSL2OvrO3SNCH4Ei/jr4YtcJByNBECNPIUAAAAAuRcBAAAPtwQvA/2LbCQQZjvBD7/Ii0U8fwwPtoRIAP7//2oH6wcPtkQBqGoIUFNV6CkDAAAPt0cCZoXAdA0rdwSYUFZTVegTAwAAi0U4" & _
    "g87/i1wkIL8eAAAAi+iL/40EN5krwtH4jQxAO1yNBH0Ei/jr6ztcjQh+BIvw6+GLfCQcjQRAjTSFAAAAAAP1i2wkEGoFD78Oi0U8D7YEyFBXVei5AgAAD7dGAmaFwHQPmFCLwytGBFBXVeihAgAAi0QkJIXAD4/Z/v//X15bXYPECMIMAMzMzMzMzMzMzMzMU1WLbCQMVldqCItFKP/Qi1wkKIv4i8u4AQAAANPgvgEAAACLTCQsSNPmiUQkKI0E9QAAAABQi0Uo/9CJRwQzyY1G/4PK/4kHhcB4I41kJACLRwRmiVTIAotHBMYEyACLRwTHRMgEAAAAAEE7D37hM+05bCQgfmmLdCQc6wONSQAPtgQuO8N+UYtEJBiLBKiLyCNMJCg7TCQkdT6Ly9P4Iwc7B380i08EZolswQKLTwQPthQuK9ONNMEPtg47yn0CiBaLdCQcugEAAAAPtgwuK8vT4gPCOwd+zEU7bCQgfKAz9jk3fGeLVCQsi2wkHIv/" & _
    "i0cEjQzwD7YBO8J+S4PI/2aJQQKLRwSNDPAPtgErwoP4B34FuAcAAABQjQQTiBFQi8uLxtPgC0QkLFD/dCQsVf90JCz/dCQs6MP+//+LTwSLVCQsiUTxBEY7N36ji8dfXl1bwhwAzMzMzMzMzMzMzIHsBAUAADPAM9JVi6wkFAUAAIlUJASJRCQMiUQkEIlEJBSJRCQYiUQkHIlEJCCJRCQkiUQkKIlEJCyJRCQwiUQkNIlEJDiJRCQ8iUQkQIlEJERWi7QkFAUAAIXtfhmL/w+2DDD/RIwMO9F9AovRQDvFfO2JVCQIM9IzyesDjUkAiVQMUANUDBCDwQQD0oP5PHzuVzP/he1+SVONnCSUAAAAD7Y0N8cDAAAAAItUtFSNQgGJRLRUhfZ0GDPAjWQkAIvKA8CD4QHR+gvBg+4BdfCJA4u0JBwFAABHg8MEO/18wFuLRCQMX4P4CXwFuAkAAABQagBqAFX/tCQkBQAAjYQkoAAAAFCLhCQoBQAA/zDo" & _
    "jP3//15dgcQEBQAAwgwAzItEJAxWi3QkDItOENPgCUYMi0QkFAPBiUYQg/gIfDlXi3wkDItGCDlGBHwPA8BQ/zaJRgiLRyz/0IkGi1YEiw6KRgyIBAr/RgTBbgwIg0YQ+IN+EAh9zV9ewhAAVYvsiwCAQMAgoGDgEJBQ0DCwcPAIiEjIKKho6BiYWNg4uHj4BIRExCSkZOQUlFTUNLR09AyMTMwsrGzsHJxc3Dy8fPwCgkLCIqJi4hKSUtIysnLyCopKyiqqauoamlraOrp6+gaGRsYmpmbmFpZW1ja2dvYOjk7OLq5u7h6eXt4+vn7+AYFBwSGhYeERkVHRMbFx8QmJSckpqWnpGZlZ2Tm5efkFhUXFJaVl5RWVVdU1tXX1DY1NzS2tbe0dnV3dPb19/QODQ8Mjo2PjE5NT0zOzc/MLi0vLK6tr6xubW9s7u3v7B4dHxyenZ+cXl1fXN7d39w+PT88vr2/vH59f3z+/f/8BAQAAAwAAAAMAAAACAQAA" & _
    "BAAAAAQAAAADAQAABQAAAAUAAAAEAQAABgAAAAYAAAAFAQAABwAAAAcAAAAGAQAACAAAAAgAAAAHAQAACQAAAAkAAAAIAQAACgAAAAoAAAAJAQEACwAAAAwAAAAKAQEADQAAAA4AAAALAQEADwAAABAAAAAMAQEAEQAAABIAAAANAQIAEwAAABYAAAAOAQIAFwAAABoAAAAPAQIAGwAAAB4AAAAQAQIAHwAAACIAAAARAQMAIwAAACoAAAASAQMAKwAAADIAAAATAQMAMwAAADoAAAAUAQMAOwAAAEIAAAAVAQQAQwAAAFIAAAAWAQQAUwAAAGIAAAAXAQQAYwAAAHIAAAAYAQQAcwAAAIIAAAAZAQUAgwAAAKIAAAAaAQUAowAAAMIAAAAbAQUAwwAAAOIAAAAcAQUA4wAAAAEBAAAdAQAAAgEAAAIBAAAQERIACAcJBgoFCwQMAw0CDgEPAAAAAAABAAAAAQAAAAEAAAACAAAAAgAAAAIAAAADAAAA" & _
    "AwAAAAMAAAAEAAAABAAAAAQAAQAFAAAABgAAAAUAAQAHAAAACAAAAAYAAgAJAAAADAAAAAcAAgANAAAAEAAAAAgAAwARAAAAGAAAAAkAAwAZAAAAIAAAAAoABAAhAAAAMAAAAAsABAAxAAAAQAAAAAwABQBBAAAAYAAAAA0ABQBhAAAAgAAAAA4ABgCBAAAAwAAAAA8ABgDBAAAAAAEAABAABwABAQAAgAEAABEABwCBAQAAAAIAABIACAABAgAAAAMAABMACAABAwAAAAQAABQACQABBAAAAAYAABUACQABBgAAAAgAABYACgABCAAAAAwAABcACgABDAAAABAAABgACwABEAAAABgAABkACwABGAAAACAAABoADAABIAAAADAAABsADAABMAAAAEAAABwADQABQAAAAGAAAB0ADQABYAAAAIAAAA=="
Private Const STR_THUNK_OFFSETS As String = "592|656|704|784|800|912|0|288|368|416|0|0|0|6896|7264|6640|7244"
Private Const STR_THUNK_BUILDDATE As String = "12.1.2018 17:15:52"
' end of generated code

'=========================================================================
' API
'=========================================================================

Private Const VT_BYREF                      As Long = &H4000
'--- for CreateFile
Private Const CREATE_ALWAYS                 As Long = 2
Private Const OPEN_EXISTING                 As Long = &H3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for CryptBinaryToString/CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for CNG
Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"
Private Const BCRYPT_AES_ALGORITHM          As String = "AES"
Private Const BCRYPT_HASH_LENGTH            As String = "HashDigestLength"
Private Const BCRYPT_OBJECT_LENGTH          As String = "ObjectLength"
Private Const BCRYPT_CHAINING_MODE          As String = "ChainingMode"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_SIG_LOCAL         As Long = &H4034B50
Private Const LNG_SIG_CENTRAL       As Long = &H2014B50
Private Const LNG_SIG_END           As Long = &H6054B50
Private Const LNG_SIG_DATADESC      As Long = &H7064B50
Private Const LNG_LEN_LOCAL         As Long = 30
Private Const LNG_LEN_CENTRAL       As Long = 46
Private Const LNG_LEN_END           As Long = 22
Private Const LNG_VER_DEFAULT       As Long = 10
Private Const LNG_VER_DEFLATE       As Long = 20
Private Const LNG_VER_ZIP64         As Long = 45
Private Const LNG_METHOD_DEFLATE    As Long = 8
Private Const LNG_METHOD_STORE      As Long = 0
Private Const LNG_IOBUF_SIZE        As Long = 65536
Private Const LNG_METHOD_AES        As Long = 99
Private Const LNG_EXTRADATA_AES_HEADER As Integer = &H9901
Private Const LNG_EXTRADATA_AES_VENDOR As Integer = &H4541 '-- "AE"
Private Const LNG_EXTRADATA_AES_SIZE As Long = 11
Private Const LNG_EXTRADATA_AE_2_MAXSIZE As Long = 20
Private Const LNG_ENC_HEADER_SIZE   As Long = 12
Private Const LNG_AES_AUTHCODE_SIZE As Long = 10
Private Const LNG_AES_BLOCK_SIZE    As Long = 16
Private Const LNG_PIPE_ENDED_ERROR  As Long = 109
Private Const ERR_USER_CANCEL       As String = "User cancelled"
Private Const ERR_INIT_COMPRESSOR   As String = "Cannot init deflate compressor"
Private Const ERR_COMPRESSING       As String = "Error compressing"
Private Const ERR_EMPTY_ARCHIVE     As String = "Empty archive"
Private Const ERR_INVALID_ARCHIVE   As String = "Invalid archive"
Private Const ERR_ARCHIVE_NOT_OPEN  As String = "Archive not opened"
Private Const ERR_INVALID_LOCALHDR  As String = "Invalid local header"
Private Const ERR_UNSUPPORTED_METHOD As String = "Unsupported compression method %1"
Private Const ERR_INIT_DECOMPRESSOR As String = "Cannot init inflate decompressor"
Private Const ERR_READING_ARCHIVE   As String = "Error reading archive"
Private Const ERR_DECOMPRESSING     As String = "Error decompressing"
Private Const ERR_CRC_CHECK         As String = "CRC check failed"
Private Const ERR_ENTRY_INVALID_SIG As String = "Entry %1 has invalid signature"
Private Const ERR_ENTRY_NO_FILENAME As String = "Entry %1 has no filename"
Private Const ERR_MISSING_ECD_RECORD As String = "Missing end-of-central-directory record"
Private Const ERR_UNSUPPORTED_ENCRYPTION As String = "Unsupported encryption"
Private Const ERR_PASSWORD_REQUIRED As String = "Password required"
Private Const ERR_INVALID_PASSWORD  As String = "Invalid password"
Private Const ERR_INVALID_AUTHCODE  As String = "Invalid authentication code"
Private Const ERR_INVALID_BYTEARRAY As String = "Invalid byte array"
Private Const STR_BUFFER            As String = "[buffer]"
Private Const STR_STREAM            As String = "[stream]"
Private Const STR_ENTRY             As String = "[entry %1]"

Private m_uRtbl                 As UcsZlibRelocTableType
Private m_sLastError            As String
Private m_bCancel               As Boolean

Private Type UcsZlibRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MemNonce            As Long
    MemXor              As Long
    ZipCrypt            As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Table          As Long
End Type

Private Type UcsZlibBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
End Type

Private Enum UcsRelocIndexesEnum
    ucsIdx_CompressInit = 0
    ucsIdx_CompressCleanup
    ucsIdx_CompressBlock
    ucsIdx_DecompressInit
    ucsIdx_DecompressCleanup
    ucsIdx_DecompressBlock
    ucsIdx_CalcCrc32
    ucsIdx_MemNonce
    ucsIdx_MemXor
    ucsIdx_ZipCrypt
    ucsIdx_MallocImpl
    ucsIdx_ReallocImpl
    ucsIdx_FreeImpl
    ucsIdx_LenCodes
    ucsIdx_DistCodes
    ucsIdx_MirrorBytes
    ucsIdx_LenLenMap
End Enum

Private Enum UcsHeaderFlagsEnum
    ucsZcfEncrypted = 2 ^ 0                         ' If set, indicates that the file is encrypted
    ucsZcfCompressOptionMask = 2 ^ 1 Or 2 ^ 2       ' 0 - normal, 1 - maximum, 2 - fast, 3 - super fast
    ucsZcfHasDataDescriptor = 2 ^ 3                 ' Values are put in the data descriptor immediately following the compressed data
    ucsZcfStrongEncrypted = 2 ^ 6                   ' Strong encryption
    ucsZcfUseUtf8 = 2 ^ 11                          ' Language encoding flag (EFS)
End Enum

Private Type UcsZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacHashLen         As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 1)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    TradKey(0 To 3)     As Long
End Type

'=========================================================================
' Properties
'=========================================================================

Property Get SemVersion() As String
    SemVersion = "0.2.8"
End Property

Property Get ThunkBuildDate() As String
    ThunkBuildDate = STR_THUNK_BUILDDATE
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property



Public Function Inflate(baBuffer() As Byte, baOutput() As Byte) As Boolean
    Const FUNC_NAME     As String = "Inflate"
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
    End If
    uBuf.InBlock = VarPtr(baBuffer(0))
    uBuf.InLen = UBound(baBuffer) + 1
    lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_DECOMPRESSING
    End If
    ReDim baOutput(0 To uBuf.OutLen - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal uBuf.OutBlock, uBuf.OutLen)
    '--- success
    Inflate = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function Deflate(baBuffer() As Byte, baOutput() As Byte, Optional ByVal Level As Long = 6) As Boolean
    Const FUNC_NAME     As String = "Deflate"
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
    End If
    'uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = 16 'At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = 65 'At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
    uBuf.InBlock = VarPtr(baBuffer(0))
    uBuf.InLen = UBound(baBuffer) + 1
    uBuf.Final = 1
    lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_COMPRESSING
    End If
    ReDim baOutput(0 To uBuf.OutLen - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal uBuf.OutBlock, uBuf.OutLen)
    '--- success
    Deflate = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
        hCtx = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

'
Private Function pvArrPtr(vArray As Variant) As Long
    Dim lPtr            As Long
    
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        pvArrPtr = Peek(UnsignedAdd(lPtr, 12))
    End If
End Function

'= common ================================================================
'
Private Function pvSetError(Optional Source As String, Optional Description As String, Optional ByVal CanContinue As Boolean) As Boolean
    If StrPtr(Source) = 0 And StrPtr(Description) = 0 Then
        m_sLastError = vbNullString
        m_bCancel = False
    Else
                Debug.Print "Error in " & Source & ": " & Description, Timer
        m_sLastError = Description
        If Not CanContinue Then
            m_bCancel = True
        End If
        pvSetError = m_bCancel
    End If
End Function

'= zlib thunks ===========================================================

Private Function pvInitRelocTable(uRtbl As UcsZlibRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(ucsIdx_CompressInit)
        .CompressCleanup = lpThunk + vSplit(ucsIdx_CompressCleanup)
        .CompressBlock = lpThunk + vSplit(ucsIdx_CompressBlock)
        .DecompressInit = lpThunk + vSplit(ucsIdx_DecompressInit)
        .DecompressCleanup = lpThunk + vSplit(ucsIdx_DecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(ucsIdx_DecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(ucsIdx_CalcCrc32)
        .MemNonce = lpThunk + vSplit(ucsIdx_MemNonce)
        .MemXor = lpThunk + vSplit(ucsIdx_MemXor)
        .ZipCrypt = lpThunk + vSplit(ucsIdx_ZipCrypt)
        .MallocImpl = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(ucsIdx_LenCodes)
        .DistCodes = lpThunk + vSplit(ucsIdx_DistCodes)
        .MirrorBytes = lpThunk + vSplit(ucsIdx_MirrorBytes)
        .LenLenMap = lpThunk + vSplit(ucsIdx_LenLenMap)
        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    Dim sBuffer         As String
    
    If lpThunk = 0 Then
        sBuffer = String$(50, 0)
        Call GetEnvironmentVariable("_ZIP_THUNK_" & App.hInstance & "_" & STR_THUNK_BUILDDATE, sBuffer, Len(sBuffer) - 1)
        lpThunk = Val(sBuffer)
        If lpThunk = 0 Then
            baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2)
            lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
            Call SetEnvironmentVariable("_ZIP_THUNK_" & App.hInstance & "_" & STR_THUNK_BUILDDATE, lpThunk)
        End If
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function pvGetCrc32Table() As Long
    Static aTable()     As Long
    Static bIsInit      As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lReminder       As Long
    Dim lValue          As Long

    If Not bIsInit Then
        '--- table mem allocated: 4KB
        ReDim aTable(0 To &H3FF) As Long
        For lIdx = 0 To &H3FF
            If lIdx < &H100 Then
                lReminder = 0
                lValue = lIdx
            Else
                lReminder = aTable(lIdx - &H100)
                lValue = 0
            End If
            For lJdx = 1 To 8
                If ((lReminder Xor lValue) And 1) <> 0 Then
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF Xor &HEDB88320
                Else
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF
                End If
                lValue = (lValue And &HFFFFFFFE) \ 2 And &H7FFFFFFF
            Next
            aTable(lIdx) = lReminder
        Next
        bIsInit = True
    End If
    pvGetCrc32Table = VarPtr(aTable(0))
End Function

Public Function CalcCrc32Array(baData() As Byte) As Long
    CalcCrc32Array = -1
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baData(0)), UBound(baData) + 1, VarPtr(CalcCrc32Array))
    CalcCrc32Array = CalcCrc32Array Xor -1
End Function

Public Sub CalcCrc32Ptr(ByVal lPtr As Long, ByVal lSize As Long, lCrc32 As Long)
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), lPtr, lSize, VarPtr(lCrc32))
End Sub

'= shared ================================================================

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function

Private Function ToBase64String(baInput() As Byte) As String
    Dim lSize As Long
    If (CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, 0&, lSize) <> 0) Then
        ToBase64String = String$(lSize - 1, 0)
        Call CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64String), lSize)
    End If
End Function

'
Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function


Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function


'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
End Sub

