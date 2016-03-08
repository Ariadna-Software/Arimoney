VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmparametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de la tesorería"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12120
   Icon            =   "frmparametrosT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   12120
   Begin VB.Frame FrameOpAseguradas 
      Caption         =   "Operaciones aseguradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   120
      TabIndex        =   67
      Top             =   4920
      Width           =   11895
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   11
         Left            =   8160
         TabIndex        =   76
         Tag             =   "3|N|S|||paramtesor|FechaAsegEsFra|||"
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Desde prorroga|N|S|||paramtesor|DiasAvisoDesdeProrroga|||"
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Dias aviso falta de pago|N|S|||paramtesor|DiasMaxSiniestrohasta|||"
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Dias aviso siniestro|N|S|||paramtesor|DiasMaxSiniestroDesde|||"
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Dias aviso siniestro|N|S|||paramtesor|DiasMaxAvisoHasta|||"
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   720
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Dias aviso falta de pago|N|S|||paramtesor|DiasMaxAvisoDesde|||"
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha factura para dias asegurados"
         Height          =   255
         Index           =   3
         Left            =   8640
         TabIndex        =   77
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Aviso siniestro desde prorroga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   4680
         TabIndex        =   74
         Top             =   360
         Width           =   2580
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   3720
         TabIndex        =   73
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   2640
         TabIndex        =   72
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   1320
         TabIndex        =   71
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   70
         Top             =   720
         Width           =   465
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   10
         Left            =   2280
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias aviso siniestro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2640
         TabIndex        =   69
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Aviso falta de pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Frame FrameValDefecto 
      Caption         =   "Valores por defecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1095
      Left            =   6240
      TabIndex        =   60
      Top             =   600
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Tag             =   "1|N|N|||paramtesor|contapag|||"
         Top             =   240
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Tag             =   "2|N|N|||paramtesor|generactrpar|||"
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   2
         Left            =   3000
         TabIndex        =   9
         Tag             =   "3|N|N|||paramtesor|abonocambiado|||"
         Top             =   600
         Width           =   435
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   4
         Left            =   2520
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Asiento por pago"
         Height          =   255
         Left            =   840
         TabIndex        =   63
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Agrupar apunte banco"
         Height          =   255
         Left            =   840
         TabIndex        =   62
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Abonos cambiados"
         Height          =   255
         Left            =   3480
         TabIndex        =   61
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Talones proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Index           =   4
      Left            =   6240
      TabIndex        =   54
      Top             =   6240
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   360
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "talon proveedor|T|S|||paramtesor|talonctapro|||"
         Text            =   "0"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1800
         TabIndex        =   55
         Text            =   "Text4"
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "  "
         Height          =   225
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Tag             =   "1|N|N|||paramtesor|contatalonptepro|||"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "Cancelacion   cliente"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   57
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   5
         Left            =   2040
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Contabiliza contra cuentas puente"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   56
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   5
         Left            =   1560
         Top             =   750
         Width           =   240
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Pagarés proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   50
      Top             =   6240
      Width           =   6015
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   1800
         TabIndex        =   51
         Text            =   "Text4"
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   360
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "pagare proveedor|T|S|||paramtesor|pagarectapro|||"
         Text            =   "0"
         Top             =   720
         Width           =   1140
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   6
         Left            =   360
         TabIndex        =   22
         Tag             =   "1|N|N|||paramtesor|contapagareptepro|||"
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Cancelacion   cliente"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   53
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   6
         Left            =   1560
         Top             =   750
         Width           =   240
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   6
         Left            =   2160
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Contabiliza contra cuentas puente"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   52
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Efectos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Index           =   2
      Left            =   6240
      TabIndex        =   46
      Top             =   1680
      Width           =   5775
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   1800
         TabIndex        =   64
         Text            =   "Text4"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   360
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Cta Efecto come.|T|S|||paramtesor|ctaefectcomerciales|||"
         Text            =   "0"
         Top             =   1440
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   360
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Efecto|T|S|||paramtesor|RemesaCancelacion|||"
         Text            =   "0"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   1800
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "  "
         Height          =   225
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Tag             =   "1|N|S|||paramtesor|contaefecpte|||"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "Efectos descontados a cobrar"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   65
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   10
         Left            =   1560
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   9
         Left            =   1560
         Top             =   720
         Width           =   240
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   3
         Left            =   3960
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Efectos descontados"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   49
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Contabiliza contra cuentas de efectos"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   48
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Pagarés clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Index           =   1
      Left            =   6240
      TabIndex        =   42
      Top             =   3720
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Tag             =   "1|N|N|||paramtesor|contapagarepte|||"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   360
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Pagare cliente|T|S|||paramtesor|pagarecta|||"
         Text            =   "0"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   1800
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Contabiliza contra cuentas puente"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   45
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   2
         Left            =   2040
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   8
         Left            =   1560
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Cancelacion   cliente"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   44
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Talones cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   3720
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "  "
         Height          =   225
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Tag             =   "1|N|N|||paramtesor|contatalonpte|||"
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   1800
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   360
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Talon cliente|T|S|||paramtesor|taloncta|||"
         Text            =   "0"
         Top             =   720
         Width           =   1140
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   7
         Left            =   1560
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Contabiliza contra cuentas puente"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   26
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   0
         Left            =   2040
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Cancelacion   cliente"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   36
      Tag             =   "Codigo|N|N|0|1|paramtesor|codigo||S|"
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame66 
      Height          =   3015
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   6015
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "%Intereses credito / tarjeta|N|N|0||paramtesor|InteresesCobrosTarjeta|||"
         Top             =   2640
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   10
         Left            =   3480
         TabIndex        =   3
         Tag             =   "3|N|N|||paramtesor|Nor19xVto|||"
         Top             =   1560
         Width           =   315
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   9
         Left            =   600
         TabIndex        =   4
         Tag             =   "3|N|S|||paramtesor|EliminaRecibidosRiesgo|||"
         Top             =   1920
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Tag             =   "Respon|T|S|||paramtesor|Responsable|||"
         Top             =   2640
         Width           =   4005
      End
      Begin VB.CheckBox Check1 
         Caption         =   " "
         Height          =   225
         Index           =   8
         Left            =   600
         TabIndex        =   2
         Tag             =   "3|N|N|||paramtesor|comprobarinicio|||"
         Top             =   1560
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Partidas pendientes apliacion|T|N|||paramtesor|par_pen_apli|||"
         Top             =   1080
         Width           =   1125
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1740
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   34
         Text            =   "Text4"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Cta beneficios bancarios|T|N|||paramtesor|ctabenbanc|||"
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "% Interes tarjeta"
         Height          =   195
         Index           =   16
         Left            =   4320
         TabIndex        =   78
         Top             =   2400
         Width           =   1125
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   12
         Left            =   5520
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Norma 19 por fecha vto"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   75
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   11
         Left            =   3240
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   9
         Left            =   240
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   8
         Left            =   240
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Eliminar en recepcion de documentos al eliminar riesgo"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   66
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   7
         Left            =   1200
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Reesponsable"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   59
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Comprobar riesgo al inicio"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   58
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   1
         Left            =   2280
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Partidas pendientes aplicacion"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   2175
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   4
         Left            =   1440
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta beneficios bancarios"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   240
      MaxLength       =   10
      TabIndex        =   31
      Text            =   "1/2/3"
      Top             =   1965
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7440
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   600
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   7440
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   741
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   1140
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha inicio"
      Height          =   255
      Index           =   27
      Left            =   240
      TabIndex        =   32
      Top             =   1710
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   7440
      Width           =   2310
   End
End
Attribute VB_Name = "frmparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Dim RS As ADODB.Recordset
Dim Modo As Byte

Dim I As Integer


Private MaxLen As Integer 'Para los txt k son de ultimo nivel o de nivel anterior
                          'Ej:

Private Sub Check1_Click(Index As Integer)

'    If Modo <> 2 Then
'        If Check1(Index).Value = 1 Then
'            Check1(Index).Value = 0
'        Else
'            Check1(Index).Value = 1
'        End If
'    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 2 Then Exit Sub
    KeyPressGral KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim ModificaClaves As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 0
        'Preparao para modificar
        PonerModo 2
        
    Case 1
        
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then PonerModo 0
        End If
    
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                'CambiaPath True
                
                
                ModificaClaves = False
'                If vUsu.Nivel = 0 Then
'                    If vParam.fechaini <> CDate(Text1(0).Text) Then
'                        ModificaClaves = True
'                        cad = " fechaini = '" & Format(vParam.fechaini, FormatoFecha) & "'"
'                    End If
'                End If
                If ModificaClaves Then
                    If ModificaDesdeFormularioClaves(Me, Cad) Then
                        ReestableceVPARAM
                        PonerModo 0
                    End If
                Else
                    If ModificaDesdeFormulario(Me) Then PonerModo 0
                End If
'                CambiaPath False
            End If

    End Select
    
    'Si el modo es 0 significa k han insertado o modificado cosas
    If Modo = 0 Then _
        MsgBox "Para que los cambios tengan efecto debe reiniciar la aplicación.", vbExclamation
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub





Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
End Sub


Private Sub cmdCancelar_Click()
If Modo = 2 Then PonerCampos
PonerModo 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    '``
End Sub

Private Sub Form_Load()
    Me.Top = 200
    Me.Left = 100
    Limpiar Me
    Me.Icon = frmPpal.Icon
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 15
    End With
    CargaImagenesAyudas Me.imgCta, 1  'Lupa
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    
    
    FrameTalones(3).Visible = False
    FrameTalones(4).Visible = False
    FrameOpAseguradas.Visible = False
    
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select * from paramtesor"
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        'No hay datos
        Limpiar Me
        PonerModo 1
    Else
        If Not vParam Is Nothing Then FrameOpAseguradas.Visible = vParam.TieneOperacionesAseguradas
    
        PonerCampos
        PonerModo 0
        'Campos que nos se tocaran los ponemos con colorcitos bonitos
        If vUsu.Nivel <> 0 Then
            Text1(0).BackColor = &H80000018
            Text1(1).BackColor = &H80000018
        End If

        
    End If
    Toolbar1.Buttons(1).Enabled = (vUsu.Nivel <= 1)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    PonerLongitudCampoNivelAnterior
End Sub

'Vamos a fijar en varios campos el maxlen
'en funcion de los digitos a ultimo nivel
'
Private Sub PonerLongitudCampoNivelAnterior()
    On Error GoTo EPonerLongitudCampoNivelAnterior
    
    
    I = DigitosNivel(vEmpresa.numnivel - 1)
    If I = 0 Then I = 4
    MaxLen = I
    
  
    Exit Sub
EPonerLongitudCampoNivelAnterior:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub frmC_Selec(vFecha As Date)
    imgFec(0).Tag = vFecha
End Sub


Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Me.Tag = CadenaSeleccion
End Sub





Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
Dim C2 As String

    Select Case Index
    Case 0, 2, 3, 5, 6
        'Cancelarcion cliente
        C2 = "cuando dentro del punto ""Recepción de documentos"" se realice la contabilización"
        Select Case Index
        Case 0, 5
            C = "talones"
        Case 2, 6
            C = "pagarés"
        Case Else
            C = "Efectos"
            C2 = "cuando dentro del punto ""Cancelación cliente"" del apartado Remesas  se realice el abono de la remesa,"
        End Select
        'If Index > 0 Then C = C & " de PROVEEDORES"
        C = "Para la cancelacion de los " & C & ":" & vbCrLf & vbCrLf
        C = C & "Si tiene marcada la opcion de 'Contabiliza contra cuentas puente', " & C2
        C = C & " tendremos dos opciones:" & vbCrLf
        If Index = 3 Then C = C & "  Efectos descontados" & vbCrLf
        C = C & "    -   Una única cuenta a último nivel (Ej: 4310000), con lo que todos los apuntes irán a esa cuenta genérica." & vbCrLf
        C = C & "    -   Introducir una cuenta raíz a 4 dígitos (Ej: 4310), con lo que el programa creará cuentas a último nivel haciéndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        
        
        'Nuevo Nov 2009
        If Index = 3 Then
            C = C & vbCrLf & vbCrLf
            C = C & "    Efectos descontados a cobrar " & vbCrLf
            C = C & "    -   Una única cuenta a último nivel (Ej: 4310000), con lo que todos los apuntes irán a esa cuenta genérica." & vbCrLf
            C = C & "    -   Introducir una cuenta raíz a 4 dígitos (Ej: 4310), con lo que el programa creará cuentas a último nivel haciéndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        End If
        
    Case 1
        C = "Cuenta beneficios bancarios." & vbCrLf & vbCrLf
        C = C & "Si no esta indicada en la configuración del banco  " & vbCrLf
        C = C & "con el que estemos trabajando, utilizará esta cuenta  " & vbCrLf
    Case 4
        C = "Valores que ofertará para la contabilización de cobros/pagos. " & vbCrLf
        C = C & "Luego podrá ser modificado para cada caso  " & vbCrLf
    Case 7
        C = "Responsable para poder firmar en documentos(recibos, cheques)"
        
    Case 8
        C = "Al entrar en la empresa que compruebe si hay posibilidad de eliminar "
        C = C & vbCrLf & "riesgo, tanto en efectos como en talones y pagarés"
    Case 9
        C = "Cuando eliminamos riesgo en talones y pagarés, eliminar tambien en la tabla de  "
        C = C & vbCrLf & "recepcion de documentos."
    Case 10
        C = "Operaciones aseguradas. "
        C = C & vbCrLf & "Dias maximo(desde/hasta) para mostrar avisos de falta de pago y/o de siniestro"

        C = C & vbCrLf & "Check  'Fecha factura...'"
        C = C & vbCrLf & "Para calcular los dias de aviso,riesgo, prorroga... puede coger"
        C = C & vbCrLf & "la fecha factura o la fecha de vencimiento."
    

    Case 11
        C = "Norma 19. "
        C = C & vbCrLf & "Se contabilizara la remesa por fecha vencimiento"
        C = C & vbCrLf & "Tantos apuntes como fechas distintas haya en la remesa"
    
    Case 12
        C = "% Intereses tarjeta"
        C = C & vbCrLf & "Porcentaje anual de interes para las ventas a credito(Forpa: Tarjeta) "
        C = C & vbCrLf & "Calculo:  (% / 365) * dias_desde_vto "
        C = C & vbCrLf & vbCrLf & "Una vez impresos los recibos, si la impresión es correcta, graba columna gastos "

        '
    
    End Select
    MsgBox C, vbInformation
End Sub

Private Sub imgcta_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    Me.Tag = ""
    'Para el text de despues
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 0
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Me.Tag <> "" Then
        Text4(Index).Text = RecuperaValor(Me.Tag, 2)
        Text1(Index).Text = RecuperaValor(Me.Tag, 1)
    End If
    Me.Tag = ""
End Sub



Private Sub imgFec_Click(Index As Integer)
    Dim F As Date
    'En los tag
    'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
    F = Now
    imgFec(0).Tag = ""
    If Text1(Index).Text <> "" Then
        If IsDate(Text1(Index).Text) Then F = Text1(Index).Text
    End If
    Set frmC = New frmCal
    frmC.Fecha = F
    frmC.Show vbModal
    Set frmC = Nothing
    If imgFec(0).Tag <> "" Then
        If IsDate(imgFec(0).Tag) Then Text1(1).Text = Format(CDate(imgFec(0).Tag), "dd/mm/yyyy")
    End If
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
    
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim Cad As String
    Dim SQL As String
    Dim Valor As Currency
    
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)


    'Si queremos hacer algo ..
    Select Case Index
    Case 1
        If Text1(Index).Text = "" Then Exit Sub
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta : " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Text1(Index).SetFocus
            Exit Sub
        End If
                        
'    Case 0
'        If Text1(Index).Text = "" Then Exit Sub
'        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(Index).Text, "T")
'        If SQL = "" Then
'            MsgBox "La cuenta no existe: " & Text1(Index).Text, vbExclamation
'            Text1(Index).Text = ""
'            Text1(Index).SetFocus
'        End If
'    Case 5, 9, 24
'     ' Diarios
'       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
'       SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(Index).Text)
'       If SQL = "" Then
'            SQL = "Codigo incorrecto"
'            Text1(Index).Text = "-1"
'        End If
'       Text2(Index).Text = SQL
'    Case 6, 7, 10, 11, 23
'       'Conceptos
'       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
'       SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(Index).Text)
'       If SQL = "" Then
'            SQL = "Codigo incorrecto"
'            Text1(Index).Text = "-1"
'        End If
'       Text2(Index).Text = SQL
        '....
    Case 0, 4
        Cad = Text1(Index).Text
        If Cad = "" Then
            Text4(Index).Text = ""
            Exit Sub
        End If
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text1(Index).Text = Cad
            Text4(Index).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(Index).Text = Cad
            Text4(Index).Text = SQL
            If Modo > 2 Then Text1(Index).SetFocus
        End If
        
'    Case 5, 6
'        If Text1(Index).Text = "" Then Exit Sub
'        cad = ""
'        If Not IsNumeric(Text1(Index)) Then cad = "Campo debe ser numerico: " & Text1(Index).Text
'        If cad = "" Then
'            If Len(Text1(Index).Text) <> Text1(5).MaxLength Then cad = "Longitud debe ser " & Text1(5).MaxLength & " digitos. " & Text1(Index).Text
'        End If
'        If cad <> "" Then
'            MsgBox cad, vbExclamation
'            Text1(Index).Text = ""
'            Ponerfoco Text1(Index)
'        End If

    Case 5, 6, 7, 8, 9, 10
        Text4(Index).Text = ""
        If Text1(Index).Text = "" Then Exit Sub
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text1(Index).Text = ""
            Exit Sub
        End If
        I = Len(Text1(Index).Text)
        NumRegElim = InStr(1, Text1(Index).Text, ".")
        If NumRegElim = 0 Then
            If I <> vEmpresa.DigitosUltimoNivel And I <> MaxLen Then
                MsgBox "Longitud de campo incorrecta. Digitos: " & vEmpresa.DigitosUltimoNivel & " o " & MaxLen, vbExclamation
                Text1(Index).Text = ""
                Exit Sub
            End If
        End If
        
        'Llegados aqui, si es de ultimo nivel pondre la cuenta
        If NumRegElim > 0 Or I = vEmpresa.DigitosUltimoNivel Then
            Cad = Text1(Index).Text
            If CuentaCorrectaUltimoNivel(Cad, SQL) Then
                Text1(Index).Text = Cad
            Else
                MsgBox SQL, vbExclamation
                Text1(Index).Text = ""
                SQL = ""
            End If
        Else
            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(Index).Text, "T")
        End If
        Text4(Index).Text = SQL
    Case 11, 12, 13, 14, 16
        If Text1(Index).Text = "" Then Exit Sub
        
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text1(Index).Text = ""
            Ponerfoco Text1(Index)
        Else
            If Index = 16 Then
                If InStr(1, Text1(Index).Text, ",") > 0 Then
                    Valor = ImporteFormateado(Text1(Index).Text)
                Else
                    Valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                End If
                Text1(Index).Text = Format(Valor, FormatoImporte)
            End If
        End If
        
    End Select
    '---
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim Valor As Boolean
    Modo = Kmodo
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        Valor = True
        Label3.Caption = ""

    Case 1
        'Preparamos para que pueda insertar
        Valor = False
        Label3.Caption = "INSERTAR"
        Label3.ForeColor = vbBlue

    Case 2
        Valor = False
        Label3.Caption = "MODIFICAR"
        Label3.ForeColor = vbRed

    End Select
    cmdAceptar.Visible = Modo > 0
    cmdCancelar.Visible = Modo > 0
    'Ponemos los valores
    
    
    
    Me.Text1(0).Enabled = Not Valor
    Me.Text1(2).Enabled = Not Valor
    Me.imgCta(0).Enabled = Not Valor
    
    For I = 4 To 16
        Me.Text1(I).Enabled = Not Valor
        If I < 11 Then Me.imgCta(I).Enabled = Not Valor
    Next
    
    For I = 0 To 11
        Me.Check1(I).Enabled = Not Valor
    Next
    
       
    'Campos que solo estan habilitados para insercion
    If Not Valor Then
        Text1(0).Locked = (vUsu.Nivel > 1)
        Text1(1).Locked = (vUsu.Nivel > 1)
        
    End If
    For I = 0 To imgFec.Count - 1
        imgFec(I).Enabled = Not Text1(0).Locked
    Next I
End Sub

Private Sub PonerCampos()
    Dim Cam As String
    Dim Tabla As String
    Dim Cod As String
    
        If Adodc1.Recordset.EOF Then Exit Sub
        If PonerCamposForma(Me, Adodc1) Then

           
           'Conceptos

           
           'Cuenta de pérdidas y ganancias
           Text4(0).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(0).Text, "T")
           Text4(4).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(4).Text, "T")
           
           For I = 6 To 10
                Text4(I).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(I).Text, "T")
           Next
           
           
        End If
End Sub
'
Private Function DatosOk() As Boolean
    Dim B As Boolean
    Dim J As Integer
    Dim C As String
    
    J = 0
    C = ""
    
    If Me.Check1(3).Value = 0 Xor Text1(7).Text = "" Then C = C & "-  talones "
    If Me.Check1(4).Value = 0 Xor Text1(8).Text = "" Then C = C & "-  pagarés "
    
    
    J = Len(Text1(9).Text) + Len(Text1(10).Text)
    If Me.Check1(5).Value = 0 Xor J = 0 Then C = C & "-  efectos "
    'Proveedores
    If Me.Check1(6).Value = 0 Xor Text1(6).Text = "" Then C = C & "-  talones PROVEEDORES"
    If Me.Check1(7).Value = 0 Xor Text1(5).Text = "" Then C = C & "-  pagarés PROVEEDORES"
    
    If C = "" Then
        'Todo bien. Compruebo esto tb
        'Si pone cta puente, la primera de las ctas es obligada
        If Me.Check1(5).Value = 1 And Text1(9).Text = "" Then C = "   La cuenta puente es campo obligatorio"
    End If
    
    If C <> "" Then
        C = Mid(C, 2) 'quitamos el primer guion
        C = C & vbCrLf & "Si marca que utliza la cuenta puente entonces debe indicarla. En otro caso debe dejarla a blancos"
        MsgBox C, vbExclamation
        Exit Function
    End If
    
    
    
    
    
    C = ""
    'AHora es desde i=5
    For I = 5 To 9
        Text1(I).Text = Trim(Text1(I).Text)
        
        If Text1(I).Text <> "" Then
            
            If Len(Text1(I).Text) <> vEmpresa.DigitosUltimoNivel Then
                                        'Aqui tenemos los digitos a ultnivel-1
                If Len(Text1(I).Text) <> MaxLen Then
                    C = C & RecuperaValor(Text1(I).Tag, 1) & vbCrLf
                End If
            End If
        End If
    Next I
    
    If C <> "" Then
        C = "Error en la longitud de las cuentas para: " & vbCrLf & C
        C = C & vbCrLf & "Ha de tener longitud "
        C = C & "a " & vEmpresa.DigitosUltimoNivel & " digitos o "
        C = C & "a " & MaxLen & " digitos. "
        
        MsgBox C, vbExclamation
        Exit Function
    End If
    
    DatosOk = False
    
    'Si es nuevo
    If Modo = 1 Then Text1(3).Text = 1

    B = CompForm(Me)
    If Not B Then Exit Function

    

    DatosOk = B
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'Modificar
         PonerModo 2
         Ponerfoco Text1(0)
    Case 2
        'Salir
        Unload Me
    End Select
End Sub



Private Sub ReestableceVPARAM()
    Set vParam = Nothing
    Set vParam = New Cparametros
    vParam.Leer
End Sub
