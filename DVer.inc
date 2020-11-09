// Product Name                    Version   Conditional/Compiler Version
//

// Embarcadero Delphi 10.4 Sydney   27.0     VER340                 UNICODE
// Embarcadero Delphi 10.3 Rio      26.0     VER330                 UNICODE
// Embarcadero Delphi 10.2 Tokyo    25.0     VER320                 UNICODE
// Embarcadero Delphi 10.1 Berlin   24.0     VER310                 UNICODE
// Embarcadero Delphi 10 Seattle    23.0     VER300                 UNICODE
// Embarcadero Delphi XE 8          22.0     VER290                 UNICODE
// Embarcadero Delphi XE 7          21.0     VER280                 UNICODE
// Embarcadero Delphi XE 6          20.0     VER270                 UNICODE
// Embercadero Delphi XE 5          19.0     VER260                 UNICODE
// Embarcadero Delphi XE 4          18.0     VER250                 UNICODE
// Embarcadero Delphi XE 3          17.0     VER240                 UNICODE
// Embarcadero Delphi XE 2          16.0     VER230                 UNICODE
// Embarcadero Delphi XE            15.0     VER220                 UNICODE
// Embarcadero RAD 2010             14.0     VER210                 UNICODE
// CodeGear Delphi 2009             12.0     VER200                 UNICODE
// CodeGear C++ Builder 2009        12.0     VER200                 UNICODE
// CodeGear Delphi 2007 for .NET    11.0     VER190
// CodeGear Delphi 2007 for Win32   11.0     VER180 and VER185
// Borland Developer Studio 2006    10.0     VER180
// Borland Delphi 2005               9.0     VER170
// Borland Delphi 8 for .NET         8.0     VER160 *
// C++BuilderX                         ?     ?
// Borland C#Builder                 1.0     VER160 *
// Borland Delphi 7                  7.0     VER150
// Borland Kylix 3                   3.0     VER140 **
// Borland C++Builder 6                ?     VER140 **
// Borland Kylix 2                   2.0     VER140 **
// Borland Delphi 6                  6.0     VER140 **
// Borland Kylix                     1.0     VER140 **
// Borland C++Builder 5                ?     VER130 ***
// Borland Delphi 5                  5.0     VER130 ***
// Borland C++Builder 4                ?     VER125
// Borland Delphi 4                  4.0     VER120
// Borland C++Builder 3                ?     VER110 ****
// Borland Delphi 3                  3.0     VER100
// Borland C++ 5                       ?     ?
// Borland C++Builder 1                ?     VER93
// Borland Delphi 2                  2.0     VER90
// Borland C++ 4.5                     ?     ?
// Borland Delphi                    1.0     VER80
// Borland C++ 4                       ?     ?
// Borland Pascal 7                  7.0     VER70
// Borland C++ 3.1                     ?     ?
// Turbo Pascal for Windows 1.5      1.5     VER70
// Turbo C++ for DOS 3                 ?     ?
// Borland C++ 3                       ?     ?
// Turbo C++ for Windows 3 (Win16)     ?     ?
// Turbo Pascal for Windows 1.0      1.0     ?
// Borland C++ 2                       ?     ?
// Turbo Pascal 6                    6.0     VER60
// Turbo C++ for DOS                   ?     ?
// Turbo C for DOS 2                   ?     ?
// Turbo Pascal 5.5                  5.5     VER55
// Turbo C for DOS 1.5                 ?     ?
// Turbo Pascal 5                    5.0     VER50
// Turbo Pascal 4                    4.0     VER40
// Turbo C for DOS                     ?     ?
// Turbo Pascal 3                    3.0     ?
// Turbo Pascal 2                    2.0     ?
// Turbo Pascal 1                    1.0     ?
//
// Source : http://delphi.wikia.com/wiki/Borland_Compiler_Conditional_Defines

{$UNDEF DESKTOP}
{$UNDEF MOBILE}
{$UNDEF MAC_DESKTOP}
{$UNDEF WIN_DESKTOP}
{$UNDEF DESKTOP_MAC}
{$UNDEF DESKTOP_WIN}


{$UNDEF SYDNEY}      {$UNDEF SYDNEY_Only}
{$UNDEF RIO}         {$UNDEF RIO_Only}
{$UNDEF TOKYO}       {$UNDEF TOKYO_Only}
{$UNDEF BERLIN}      {$UNDEF BERLIN_Only}
{$UNDEF SEATTLE}     {$UNDEF SEATTLE_Only}
{$UNDEF DX}          {$UNDEF DX_Only}
{$UNDEF XE8}         {$UNDEF XE8_Only}
{$UNDEF XE7}         {$UNDEF XE7_Only}
{$UNDEF XE6}         {$UNDEF XE6_Only}
{$UNDEF XE5}         {$UNDEF XE5_Only}
{$UNDEF XE4}         {$UNDEF XE4_Only}
{$UNDEF XE3}         {$UNDEF XE3_Only}
{$UNDEF XE2}         {$UNDEF XE2_Only}
{$UNDEF XE}          {$UNDEF XE_Only}
{$UNDEF Delphi2010}  {$UNDEF Delphi2010_only}
{$UNDEF Delphi2009}  {$UNDEF Delphi2009_only}
{$UNDEF Delphi2007}  {$UNDEF Delphi2007_only}
{$UNDEF Delphi2006}  {$UNDEF Delphi2006_only}
{$UnDef Delphi6}     {$UnDef Delphi6_Only}
{$UnDef Delphi5}     {$UnDef Delphi5_Only}
{$UnDef Delphi4}     {$UnDef Delphi4_Only}
{$UnDef Delphi3}     {$UnDef Delphi3_Only}
{$UnDef Delphi2}     {$UnDef Delphi2_Only}
{$UnDef Delphi1}     {$UnDef Delphi1_Only}

{$IFDEF CONDITIONALEXPRESSIONS} // Delphi 6 or higher

  {$IF DEFINED(ANDROID)}
    {$DEFINE MOBILE}
  {$ELSEIF DEFINED(WIN32) OR DEFINED(WIN64) }
    {$DEFINE DESKTOP}
    {$DEFINE WIN_DESKTOP}
    {$DEFINE DESKTOP_WIN}
  {$ELSEIF DEFINED(MACOS)}
    {$IF DEFINED(IOS)}
      {$DEFINE MOBILE}
    {$ELSE}
      {$DEFINE DESKTOP}
      {$DEFINE MAC_DESKTOP}
      {$DEFINE DESKTOP_MAC}
    {$ENDIF}
  {$ELSE}
    {$MESSAGE ERROR 'Unsupported Deployment target'}
  {$ENDIF}

  {$IF CompilerVersion = 34.0} {$Define SYDNEY_Only}      {$ENDIF}
  {$IF CompilerVersion = 33.0} {$Define RIO_Only}         {$ENDIF}
  {$IF CompilerVersion = 32.0} {$Define TOKYO_Only}       {$ENDIF}
  {$IF CompilerVersion = 31.0} {$Define BERLIN_Only}      {$ENDIF}
  {$IF CompilerVersion = 30.0} {$Define SEATTLE_Only}     {$ENDIF}
  {$IF CompilerVersion = 29.0} {$Define XE8_Only}         {$ENDIF}
  {$IF CompilerVersion = 28.0} {$Define XE7_Only}         {$ENDIF}
  {$IF CompilerVersion = 27.0} {$Define XE6_Only}         {$ENDIF}
  {$IF CompilerVersion = 26.0} {$Define XE5_Only}         {$ENDIF}
  {$IF CompilerVersion = 25.0} {$Define XE4_Only}         {$ENDIF}
  {$IF CompilerVersion = 24.0} {$Define XE3_Only}         {$ENDIF}
  {$IF CompilerVersion = 23.0} {$Define XE2_Only}         {$ENDIF}
  {$IF CompilerVersion = 22.0} {$Define XE_Only}          {$ENDIF}
  {$IF CompilerVersion = 21.0} {$Define Delphi2010_Only}  {$ENDIF}  // Delphi 2010
  {$IF CompilerVersion = 20.0} {$Define Delphi2009_Only}  {$ENDIF}  // Delphi 2009
  {$IF CompilerVersion = 18.5} {$Define Delphi2007_Only}  {$ENDIF}  // Delphi 2007
  {$IF CompilerVersion = 18.0} {$Define Delphi2006_Only}  {$ENDIF}  // Delphi 2006
  {$IF CompilerVersion = 17.0} {$Define Delphi2005_Only}  {$ENDIF}  // Delphi 2005
  {$IF CompilerVersion = 15.0} {$Define Delphi7_Only}     {$ENDIF}  // Delphi 7
  {$IF CompilerVersion = 14.0} {$Define Delphi6_Only}     {$ENDIF}  // Delphi 6

  {$IF CompilerVersion >= 32.0} {$Define SYDNEY}          {$ENDIF}
  {$IF CompilerVersion >= 32.0} {$Define RIO}             {$ENDIF}
  {$IF CompilerVersion >= 32.0} {$Define TOKYO}           {$ENDIF}
  {$IF CompilerVersion >= 31.0} {$Define BERLIN}          {$ENDIF}
  {$IF CompilerVersion >= 30.0} {$Define SEATTLE}         {$ENDIF}
  {$IF CompilerVersion >= 29.0} {$Define XE8}             {$ENDIF}
  {$IF CompilerVersion >= 28.0} {$Define XE7}
    {$Define DynamicArraySupport}                         {$ENDIF}
  {$IF CompilerVersion >= 27.0} {$Define XE6}             {$ENDIF}
  {$IF CompilerVersion >= 26.0} {$Define XE5}             {$ENDIF}
  {$IF CompilerVersion >= 25.0} {$Define XE4}             {$ENDIF}
  {$IF CompilerVersion >= 24.0} {$Define XE3}             {$ENDIF}
  {$IF CompilerVersion >= 23.0} {$Define XE2}             {$ENDIF}
  {$IF CompilerVersion >= 22.0} {$Define XE}              {$ENDIF}
  {$IF CompilerVersion >= 21.0} {$Define Delphi2010}      {$ENDIF}  // Delphi 2010
  {$IF CompilerVersion >= 20.0} {$Define Delphi2009}      {$ENDIF}  // Delphi 2009
  {$IF CompilerVersion >= 18.5} {$Define Delphi2007}      {$ENDIF}  // Delphi 2007
  {$IF CompilerVersion >= 18.0} {$Define Delphi2006}      {$ENDIF}  // Delphi 2006
  {$IF CompilerVersion >= 17.0} {$Define Delphi2005}      {$ENDIF}  // Delphi 2005
  {$IF CompilerVersion >= 15.0} {$Define Delphi7}         {$ENDIF}  // Delphi 7
  {$IF CompilerVersion >= 14.0} {$Define Delphi6}         {$ENDIF}  // Delphi 6

  {$IF CompilerVersion > 33.0} {$Message WARN 'UNTESTED COMPILER SUPPORT'} {$ENDIF}
  {$IF CompilerVersion < 18.5} {$Message ERROR 'UNSUPPORTED COMPILER'}     {$ENDIF}

  {$DEFINE DELPHI5}
  {$DEFINE DELPHI4}
  {$DEFINE DELPHI3}
  {$DEFINE DELPHI2}
  {$DEFINE DELPHI1}
{$ENDIF CONDITIONALEXPRESSIONS}

{$IFNDEF CONDITIONALEXPRESSIONS} // Delphi 5 or lower
  {$IfDef Ver130} // Delphi 5
    {$Define Delphi5_Only}
    {$Define Delphi5}
    {$Define Delphi4}
    {$Define Delphi3}
    {$Define Delphi2}
    {$Define Delphi1}
  {$ENDIF Ver130}

  {$IfDef Ver120} // Delphi 4
    {$Define Delphi4_Only}
    {$Define Delphi4}
    {$Define Delphi3}
    {$Define Delphi2}
    {$Define Delphi1}
  {$ENDIF Ver120}

  {$IfDef Ver100} // Delphi 3
    {$Define Delphi3_Only}
    {$Define Delphi3}
    {$Define Delphi2}
    {$Define Delphi1}
  {$ENDIF Ver100}

  {$IfDef Ver90}  // Delphi 2
    {$Define Delphi2_Only}
    {$Define Delphi2}
    {$Define Delphi1}
  {$ENDIF Ver90}

  {$IfDef Ver80}  // Delphi 1
    {$Define Delphi1_Only}
    {$Define Delphi1}
  {$ENDIF Ver80}

{$ENDIF CONDITIONALEXPRESSIONS}

{$DEFINE LittleEndian}

(*
{$IF defined(SYDNEY_Only)}
  {$MESSAGE HINT 'SYDNEY found'}
{$ELSEIF defined(RIO_Only)}
  {$MESSAGE HINT 'RIO found'}
{$ELSEIF defined(TOKYO_Only)}
  {$MESSAGE HINT 'TOKYO found'}
{$ELSEIF defined(BERLIN_Only)}
  {$MESSAGE HINT 'BERLIN found'}
{$ELSEIF defined(SEATTLE_Only)}
  {$MESSAGE HINT 'SEATTLE found'}
{$ELSEIF defined(XE8_Only)}
  {$MESSAGE HINT 'XE8 found'}
{$ELSEIF defined(XE7_Only)}
  {$MESSAGE HINT 'XE7 found'}
{$ELSEIF defined(XE6_Only)}
  {$MESSAGE HINT 'XE6 found'}
{$ELSEIF defined(XE5_Only)}
  {$MESSAGE HINT 'XE5 found'}
{$ELSEIF defined(XE4_Only)}
  {$MESSAGE HINT 'XE4 found'}
{$ELSEIF defined(XE3_Only)}
  {$MESSAGE HINT 'XE3 found'}
{$ELSEIF defined(XE2_Only)}
  {$MESSAGE HINT 'XE2 found'}
{$ELSEIF defined(XE_Only)}
  {$MESSAGE HINT 'XE found'}
{$ELSEIF defined(Delphi2010_Only)}
  {$MESSAGE HINT 'Delphi2010 found'}
{$ELSEIF defined(Delphi2009_Only)}
  {$MESSAGE HINT 'Delphi2009 found'}
{$ELSEIF defined(Delphi2007_Only)}
  {$MESSAGE HINT 'Delphi2007 found'}
{$ELSEIF defined(Delphi2006_Only)}
  {$MESSAGE HINT 'Delphi2006 found'}
{$ELSEIF defined(Delphi2005_Only)}
  {$MESSAGE HINT 'Delphi2005 found'}
{$ELSEIF defined(Delphi7_Only)}
  {$MESSAGE HINT 'Delphi7 found'}
{$ELSEIF defined(Delphi6_Only)}
  {$MESSAGE HINT 'Delphi6 found'}
{$ELSEIF}
  {$MESSAGE Error 'Delphi version unknown'}
{$ENDIF}
*)
