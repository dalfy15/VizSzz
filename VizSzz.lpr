program VizSzz;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, m_VizSzz, m_myFs, m_type, me_exl, datetimectrls, pexpandpanels { you can add units after this };

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Scaled := True;
  Application.Initialize;

  myFs := TmyFs.Create(nil);
  myFs.Hide;
  Application.CreateForm(TFVizSzz, FVizSzz);
  Application.Run;
end.


