let imageCompatPane;

function OnShowImageCompatPane() {
  if (!imageCompatPane) {
    imageCompatPane = Application.CreateTaskPane("http://localhost:3000/index.html");
    imageCompatPane.DockPosition = Application.Enum.JSKsoEnum_msoCTPDockPositionRight;
    imageCompatPane.Width = 360;
  }
  imageCompatPane.Visible = true;
}

function GetImageCompatEnabled() {
  return Boolean(Application && Application.ActiveWorkbook);
}
