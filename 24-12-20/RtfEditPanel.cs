using DevExpress.Utils.Commands;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraRichEdit.Export;
using static .Controls.Items.TransparentDocumentServer;

namespace .Panels {
    public partial class RtfEditPanel : EditGroupPanel {

        #region events
        public event EventHandler<AnimateBoxEventArg> ItemRemoved;
        public event EventHandler<List<byte[]>> RtfChanged;
        public event EventHandler<int> VerticalAlignmentChanged;
        #endregion


        #region variables
        // RichEditControl
        private TextBarAction[] Paragraphs = new TextBarAction[] {
            TextBarAction.HorizontalLeft, TextBarAction.HorizontalCenter, TextBarAction.HorizontalRight,
            TextBarAction.VerticalBottom, TextBarAction.VerticalCenter, TextBarAction.VerticalTop,
        };

        private TextBarAction[] CharactersActions = new TextBarAction[] {
            TextBarAction.Name, TextBarAction.Size, TextBarAction.ForeColor, TextBarAction.BackColor,
            TextBarAction.Bold, TextBarAction.Italic, TextBarAction.Underline
        };

        // 이미지
        private int _width = 192;
        private int _height = 128;

        // rtf
        //private PrintingSystem _printSystem = new PrintingSystem();
        //private PrintableComponentLinkItem _printLinkItem;
        private RichEditControl _printRichEdit = new RichEditControl();
        private PrintingSystem _tmpPrintSystem = new PrintingSystem();
        private PrintableComponentLinkItem _tmpPrintLinkItem;
        private PrintableComponentLink pLink;
        private RichEditControl _tmpRichEdit = new RichEditControl();
        private object rtfLock = new object();
        private TableCellVerticalAlignment vertAct = TableCellVerticalAlignment.Top;
        //private Color brickColor = Color.Black;
        private ImageExportOptions ieo;
        private RtfDocumentExporterOptions rtfOpt;
        private bool isExternalOrder = false;
        private int brickColor;
        private bool isUserAction = true;
        private TableCellInfo CurrCellInfo;
        private bool RichControlSizeChanged = false;
        private bool isDocumentLoaded = false;

#if DEBUG
        // temp
        private Stopwatch sw = new Stopwatch();
        private long timestamp = 0L;
        private long currstamp = 0L;
#endif
#endregion


        #region properties
        #endregion


        #region constuctor
        #endregion


        #region override
        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
        }
        #endregion


        #region public m

        public void RtbRtfChanged(bool isReset = false) {
            try {
                richEdit.ContentChanged -= RichEdit_ContentChanged;
                lock (rtfLock) {
#if DEBUG
                    sw.Start();
#endif
                    /// -> 필수조건 테스트
                    /*
                    UI에서 보이는 RichEditControl(이하 A)
                    ---- RTF --->
                    미리보기의 너비 높이로 설정된 RichEditControl(이하 B)
                    ---- 페이지 별 RTF List --->
                    Table[B 에서 출력되는 이미지의 수, 1]로 테이블을 생성하고 각 cell의 너비 높이도 B와 동일시하게 변경.
                    Table[i, 1]에 B의 Rtf List 아이템 반복 삽입
                    ---- Table의 각 Cell RTF --->
                    
                    */

                    _printRichEdit.RtfText = richEdit.RtfText;
                    _printRichEdit.ActiveView.BackColor = Color.FromArgb(BrickColor);

                    // 나중에 caretPos로 선택한 텍스트를 포함하는 페이지를 미리보기에 반영할 수 있게 - update
                    /*int caretPosition = richEdit.Document.CaretPosition.ToInt();
                    var dp = _printRichEdit.Document.CreatePosition(caretPosition);
                    _printRichEdit.Document.CaretPosition = dp;*/

                    // 유저단 richbox --- rtf ---> 미리보기 사이즈 richbox -req
                    SetRichEditSize(_printRichEdit, PreWidth, PreHeight);
                    // width/height 지정 후 margin 안 정해주면 에러나는듯 - req -> 더 알아볼것
                    SetRichControlMargin(_printRichEdit, 0, 0, 0, 0);

                    // rtf를 page별로 가져온 후 cell에 각 rtf를 대입하고 수직정렬.
                    // 수직정렬한 cell의 rtf를 다시 PrintComponentLink에 담기위해
                    // RichEditDocumentServer에 RTF 이식 후 이미지 출력

                    // 테이블 cell마다 부여할 rtf로 분리 - req
                    int prtRtbCnt = _printRichEdit.DocumentLayout.GetPageCount();
                    List<string> rtfList = new List<string>();
                    List<TableContentIndex> horAlignList = new List<TableContentIndex>();

                    // 마지막줄에 대한 rtf값까지 가져오는 options 추가 - rtfOpt - req

                    ///// page상관없이 paraphrase 별로 수평정렬하는 방법?
                    /////var _printRichPP = _printRichEdit.Document.Paragraphs;
                    /////List<DocumentRange> _printDocRange = _printRichPP.Select(i => i.Range).ToList();

                    // 텍스트의 정렬 정보는 RichEditcontrol의 Document에서 직접 확인해야 함
                    // 문서 내 특정 범위의 Paragraph 객체를 통해 정렬 속성 확인 가능
#if DEBUG
                    timestamp = 0;
                    Debug.WriteLine($"==============================================");
                    currstamp = sw.ElapsedMilliseconds;
                    Debug.WriteLine($"[ Start] {currstamp} \t Add RTF List.");
#endif

                    int totalPos = 0;
                    int prevPos = 0;
                    // 미리보기 사이즈로 만든 richEditControl로부터 페이지별 rtf를 리스트로 저장.
                    for (int i = 0; i < prtRtbCnt; i++) {
                        var layoutPage = _printRichEdit.DocumentLayout.GetPage(i);
                        if (layoutPage == null) break;
                        FixedRange contentRange = layoutPage.MainContentRange;
                        DocumentRange mainBodyRange = _printRichEdit.Document.CreateRange(contentRange.Start, contentRange.Length);
                        int k = 0;
                        while ( totalPos < mainBodyRange.End.ToInt()) {
                            var document = _printRichEdit.Document;
                            
                            // 같은 문단이라고 해도 페이지가 달라지면 수평 정렬 리스트를 추가할 인덱스가 달라져야 함.
                            DocumentPosition dp = document.CreatePosition(totalPos);
                            DocumentRange dpRange = _printRichEdit.Document.Paragraphs.Get(dp).Range;
                            var startIdx = dpRange.Start.ToInt();
                            var endIdx = dpRange.End.ToInt();

                            int tmpEndIdx = mainBodyRange.End.ToInt() < endIdx ? mainBodyRange.End.ToInt() : endIdx;
                            var tmpAlignment = _printRichEdit.Document.Paragraphs.Get(dp).Alignment;

                            horAlignList.Add(new TableContentIndex {
                                PpAlignment = tmpAlignment,
                                Page = i,
                                PageRow = k,
                                ParagraphPos = startIdx - prevPos < 0 ? 0 : startIdx - prevPos
                            });

                            k++;
                            totalPos = tmpEndIdx;
                        }

                        prevPos = totalPos;

                        rtfList.Add(_printRichEdit.Document.GetRtfText(mainBodyRange, rtfOpt));
                    }
#if DEBUG
                    timestamp = sw.ElapsedMilliseconds;
                    Debug.WriteLine($"[    ..] {timestamp - currstamp} \t Add RTF List.");
                    Debug.WriteLine($"[Finish] {timestamp} \t Add RTF List.");
#endif
                    /*for (int i = 0; i < horAlignList.Count;i ++) {
                        Console.WriteLine($"page : {horAlignList[i].Page}");
                        Console.WriteLine($"pRow : {horAlignList[i].PageRow}");
                        Console.WriteLine($"pIDx : {horAlignList[i].ParagraphPos}");
                        Console.WriteLine($"pAln : {horAlignList[i].PpAlignment}");
                        Console.WriteLine($"--");
                    }
                    Console.WriteLine($"list end");*/

                    // rtf담을 테이블을 포함한 richEdit doc만들고 조정 - req
                    _tmpRichEdit.Document.Delete(_tmpRichEdit.Document.Range);
                    SetRichEditSize(_tmpRichEdit, PreWidth, PreHeight);
                    SetRichControlMargin(_tmpRichEdit, 0, 0, 0, 0);

                    var brickColor = Color.FromArgb(BrickColor);
                    if (IsBrickAlphaZero())
                        brickColor = Color.Empty;

                    var tmpDoc = _tmpRichEdit.Document;
#if DEBUG
                    Debug.WriteLine($"==============================================");
                    currstamp = sw.ElapsedMilliseconds;
                    Debug.WriteLine($"[ Start] {currstamp} \t Create RTF Table.");
#endif

                    // table 설정, update 시작, 테이블 사이즈조정 - req
                    Table tmpTable = tmpDoc.Tables.Create(tmpDoc.Range.End, rtfList.Count, 1); // table = rtfList.Count행 * 1열
                    tmpTable.BeginUpdate();
                    tmpTable.SetPreferredWidth(GetUnitSize(PreWidth, _tmpRichEdit.DpiX), WidthType.Fixed);

                    tmpTable.TableBackgroundColor = brickColor;
                    float unitWidth = GetUnitSize(PreWidth, _tmpRichEdit.DpiX);
                    float unitHeight = GetUnitSize(PreHeight, _tmpRichEdit.DpiY);

                    CurrCellInfo = new TableCellInfo() { 
                        BrickColor = Color.FromArgb(BrickColor),
                        UnitWidth = GetUnitSize(PreWidth, _tmpRichEdit.DpiX),
                        UnitHeight = GetUnitSize(PreHeight, _tmpRichEdit.DpiY),
                        CellVAlign = vertAct
                    };



#if false
                    // 각 페이지가 되는 table.row.cell의 옵션 설정
                    foreach (var tmpRow in tmpTable.Rows) {
                        ///// 알파값을 0으로 설정을 못해줌 > pdf로 내보내서 정확하게 이미지 처리를 하면? >> 나중에 시도 // 우선은 Color.Black으로, 
                        tmpRow.FirstCell.BackgroundColor = brickColor;

                        // 너비 설정 - req
                        tmpRow.FirstCell.PreferredWidthType = WidthType.Fixed;
                        tmpRow.FirstCell.PreferredWidth = unitWidth;

                        // 높이 설정 - req
                        tmpRow.FirstCell.HeightType = HeightType.Exact;
                        tmpRow.FirstCell.Height = unitHeight;

                        // 수직정렬 - req
                        tmpRow.FirstCell.VerticalAlignment = vertAct;

                        // padding - req
                        tmpRow.FirstCell.LeftPadding = 0;
                        tmpRow.FirstCell.TopPadding = 0;
                        tmpRow.FirstCell.RightPadding = 0;
                        tmpRow.FirstCell.BottomPadding = 0;

                        // border line thickness - req
                        tmpRow.FirstCell.Borders.Left.LineThickness = 0;
                        tmpRow.FirstCell.Borders.Top.LineThickness = 0;
                        tmpRow.FirstCell.Borders.Right.LineThickness = 0;
                        tmpRow.FirstCell.Borders.Bottom.LineThickness = 0;
                    }
#else 
                    tmpTable.ForEachCell(new TableCellProcessorDelegate(MakeMultiplicationCell));
#endif
                    tmpTable.EndUpdate();
#if DEBUG
                    timestamp = sw.ElapsedMilliseconds;
                    Debug.WriteLine($"[    ..] {timestamp - currstamp} \t Create RTF Table. ");
                    Debug.WriteLine($"[Finish] {timestamp} \t Create RTF Table. ");
#endif


                    // 수직 정렬된 cell의 rtf를 RichEditDocumentServer로 이식 후 각 cell - req
                    // table 생성후 페이지 수 대로 row 생성한 후 각 row.firstcell에 rtf를 이식
                    using (CustomRichEditDocumentServer tmpServer = new CustomRichEditDocumentServer()) {
#if DEBUG
                        Debug.WriteLine($"==============================================");
                        currstamp = sw.ElapsedMilliseconds;
                        Debug.WriteLine($"[ Start] {currstamp} \t Pass RTF Data to table.");
#endif
                        tmpServer.Document.SetPageBackground(brickColor, true);
                        tmpServer.Options.Printing.EnablePageBackgroundOnPrint = true;

                        var tmpSections = tmpServer.Document.Sections;
                        foreach (var tmpSection in tmpSections) {
                            tmpSection.Page.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.Custom;
                            tmpSection.Page.Width = GetUnitSize(PreWidth, tmpServer.DpiX);
                            tmpSection.Page.Height = GetUnitSize(PreHeight, tmpServer.DpiY);
                            tmpSection.Margins.Left = 0;
                            tmpSection.Margins.Top = 0;
                            tmpSection.Margins.Right = 0;
                            tmpSection.Margins.Bottom = 0;
                        }

                        for (int i = 0; i < prtRtbCnt; i++) {
                            // 각 cell의 rtf를 임시로 담는 과정 - req
                            tmpServer.RtfText = rtfList[i];

                            var rowList = horAlignList.Where(j => j.Page == i).ToList();
                            for (int  j = 0; j < rowList.Count; j++) {
                                var horAlignItem = horAlignList.Where(p => p.Page == i && p.PageRow == j).FirstOrDefault();
                                if (horAlignItem != null) {
                                    //var idx = horAlignItem.ParagraphPos;
                                    var contentAlignment = horAlignItem.PpAlignment;

                                    var tmpServerDoc = tmpServer.Document;
                                    DocumentPosition dp = tmpServerDoc.CreatePosition(horAlignItem.ParagraphPos);
                                    tmpServerDoc.Paragraphs.Get(dp).Alignment = contentAlignment;
                                }
                            }

                            // 전체idx를 계산하며 페이지마다 분리를 해야 한다.
#if true
                            // cell에 rtf 적용시 알 수 없는 rtf(아마 null rtf) 가 마지막 라인에 추가됨 - req
                            var dRange = tmpDoc.InsertDocumentContent(tmpTable[i, 0].Range.Start, tmpServer.Document.Range, InsertOptions.KeepSourceFormatting);
                            // 마지막 rtf에 대한 내용 삭제 - req
                            tmpDoc.Delete(tmpDoc.CreateRange(dRange.End.ToInt() - 1, 1));
#else
                            ///// Document.AppendDocumentContent // Document.AppendRtfText 를 사용했을때 Cell의 Rtf 설정하는 방법?
#endif
                        }
                        
#if DEBUG
                        timestamp = sw.ElapsedMilliseconds;
                        Debug.WriteLine($"[    ..] {timestamp - currstamp} \t Pass RTF Data to table.");
                        Debug.WriteLine($"[Finish] {timestamp} \t Pass RTF Data to table.");
                        Debug.WriteLine($"==============================================");
                        currstamp = sw.ElapsedMilliseconds;
                        Debug.WriteLine($"[ Start] {currstamp} \t PrintableComponentLink do ExportToImage.");
#endif
                        // 위에서 생성된 table 내용의 rtf를 tmpServer.exportToImage로 내보내기 위해 옮겨담기 - req
                        tmpServer.RtfText = _tmpRichEdit.RtfText;

                        //DocumentRange mainBodyRange = _printRichEdit.Document.CreateRange(contentRange.Start, contentRange.Length);

                        // pLink.ExportToImage()로 페이지 별 이미지 스트림 바이트 내보내기 - req
                        // printableComponentLink 재사용 처리는 나중에 -> PrintableComponentLink 재사용 시 글자 깨지는 이슈 있음
                        pLink = new PrintableComponentLink(new PrintingSystem());
                        pLink.Component = tmpServer;
                        pLink.Margins.Left = 0;
                        pLink.Margins.Top = 0;
                        pLink.Margins.Right = 0;
                        pLink.Margins.Bottom = 0;
                        tmpServer.Document.SetPageBackground(brickColor, true);
                        tmpServer.Options.Printing.EnablePageBackgroundOnPrint = true;

                        List<byte[]> _lstBytes = new List<byte[]>();
                        int pageCnt = _tmpRichEdit?.DocumentLayout.GetPageCount() ?? 0;
                        ieo.RetainBackgroundTransparency = IsBrickAlphaZero();
                        for (int i = 0; i < pageCnt; i++) {
                            // 마지막 빈 carriage return 들어가는 경우로 페이지 늘어나는 상황 발생 
                            if (i >= pageCnt - 1) break;
                            using (var ms = new MemoryStream()) {
                                ieo.PageRange = (i + 1).ToString();
                                pLink.ExportToImage(ms, ieo);
                                _lstBytes.Add(ms.ToArray());
                            }
                        }
                        RtfChanged?.Invoke(this, _lstBytes);
#if DEBUG
                        timestamp = sw.ElapsedMilliseconds;
                        Debug.WriteLine($"[    ..] {timestamp - currstamp} \t PrintableComponentLink do ExportToImage.");
                        Debug.WriteLine($"[Finish] {timestamp} \t PrintableComponentLink do ExportToImage.");

                        sw.Reset();

                        Debug.WriteLine($"//");
#endif
                    }
                }
            }
            catch {
                Console.WriteLine($"tfChanged error");
            }
        }

        public void SetRichControlMargin(RichEditControl control, int l = 0, int t = 0, int r = 0, int b = 0) {
            // page 내부 margin 설정
            foreach (var section in control.Document.Sections) {
                section.Margins.Left = l;
                section.Margins.Top = t;
                section.Margins.Right = r;
                section.Margins.Bottom = b;
            }
        }

        public void AddRichEditControlEvt() {
            if(eItem_cbxAutoSync.Checked) { 
                // rtfTextChanged -> 내부적요소 서포트를 위한 이벤트이며, 사용자에서 사용하는 용도 아님 -> contentchanged로 변경
                richEdit.ContentChanged += RichEdit_ContentChanged;
            }
        }

        private void RichEdit_ContentChanged(object sender, EventArgs e) {
            if (richEdit.Modified && !RichControlSizeChanged && IsDocumentLoaded)
                RtbRtfChanged();
            else
                RichControlSizeChanged = false;
        }


#endregion


#region protected m
        
#endregion


#region private m
        private void InitializeEvent() {
            

            richEdit.SelectionChanged += (s, e) => {
                // 드래그 된 상태라면 폰트정보를 세팅하지 않음
                var selectedRange = richEdit.Document.Selection;
                if (selectedRange.Length > 0) return;

                // 1글자면 폰트 정보를 ui에 동기화
                DocumentPosition docPos = richEdit.Document.CaretPosition;
                DocumentRange range = richEdit.Document.CreateRange(docPos, 1);
                CharacterProperties props = richEdit.Document.BeginUpdateCharacters(range);

                // font value props
                if (props.FontName != null || props.FontName != string.Empty)
                    eItem_cbxFontName.EditValue = props.FontName;
                if ((props.FontSize ?? -1) != -1)
                    eItem_cbxFontSize.EditValue = (int)props.FontSize;

                // font style props
                eItem_btnStyleBold.Down = props.Bold ?? false;
                eItem_btnStyleItalic.Down = props.Italic ?? false;
                eItem_btnStyleUnderline.Down = props.Underline == UnderlineType.Single;

                richEdit.Document.EndUpdateCharacters(props);

                // paragraph alignment
                ParagraphProperties pp = richEdit.Document.BeginUpdateParagraphs(range);
                ParagraphAlignment horAlignment = pp.Alignment ?? ParagraphAlignment.Left;
                TextBarAction horAction = horAlignment == ParagraphAlignment.Right ? TextBarAction.HorizontalRight : horAlignment == ParagraphAlignment.Center ? TextBarAction.HorizontalCenter : TextBarAction.HorizontalLeft;
                richEdit.Document.EndUpdateParagraphs(pp);

                TextBarAction verAction = vertAct == TableCellVerticalAlignment.Bottom ? TextBarAction.VerticalBottom : vertAct == TableCellVerticalAlignment.Center ? TextBarAction.VerticalCenter : TextBarAction.VerticalTop;

                ToggleHorizontalAlignment(horAction);
                ToggleVerticalAlignment(verAction);
            };
        }

        private void InitializeValue() {
            ieo = new ImageExportOptions() {
                ExportMode = DevExpress.XtraPrinting.ImageExportMode.SingleFilePageByPage,
                PageBorderWidth = 0,
                PageBorderColor = Color.Transparent,
                RetainBackgroundTransparency = true,
                Format = DevExpress.Drawing.DXImageFormat.Png
            };

            rtfOpt = new RtfDocumentExporterOptions() {
                ExportFinalParagraphMark = DevExpress.XtraRichEdit.Export.Rtf.ExportFinalParagraphMark.Always,
            };
        }

        private void MakeMultiplicationCell(TableCell cell, int i, int j) {
            var cellInfo = CurrCellInfo;
            cell.BackgroundColor = cellInfo.BrickColor;

            cell.HeightType = HeightType.Exact;
            cell.Height = cellInfo.UnitHeight;

            cell.VerticalAlignment = cellInfo.CellVAlign;

            cell.TopPadding = 0;
            cell.LeftPadding = 0;
            cell.TopPadding = 0;
            cell.RightPadding = 0;

            cell.Borders.Left.LineThickness = 0;
            cell.Borders.Top.LineThickness = 0;
            cell.Borders.Right.LineThickness = 0;
            cell.Borders.Bottom.LineThickness = 0;
        }

        private bool IsBrickAlphaZero() {
            return Color.FromArgb(BrickColor).A == 0;
        }

        private void ModifySync() {
            RemoveRicheditControlEvt();
            var isAuto = eItem_cbxAutoSync.Checked;
            eItem_btnApplyTyping.Visibility = isAuto ? DevExpress.XtraBars.BarItemVisibility.Never : DevExpress.XtraBars.BarItemVisibility.Always;
            if (isAuto) AddRichEditControlEvt();
            
        }

        private void ModifyVerProp(TextBarAction action) {
            ModifyFont(action, true);
            ///// 수직정렬 값 변경될 때 rtf와 함께 저장되어야 함.
            int eventVal =
                vertAct == TableCellVerticalAlignment.Bottom ? (int)TableCellVerticalAlignment.Bottom :
                vertAct == TableCellVerticalAlignment.Center ? (int)TableCellVerticalAlignment.Center :
                (int)TableCellVerticalAlignment.Top;

            VerticalAlignmentChanged?.Invoke(this, eventVal);
        }

        private void SetVertProp(TableCellVerticalAlignment verAlign) {
            if (vertAct != verAlign) vertAct = verAlign;
        }

        private void PanelSizeChanged() {
            float widthInDocuments = DevExpress.Office.Utils.Units.PixelsToDocumentsF(richEdit.ClientSize.Width, richEdit.DpiX);
            float h = DevExpress.Office.Utils.Units.PixelsToDocumentsF(3000, richEdit.DpiY);

            richEdit.Document.Sections[0].Page.Width = widthInDocuments;
            richEdit.Document.Sections[0].Page.Height = h;
        }

        private void ModifyFont(TextBarAction action, bool isClicked = false) {
            var document = richEdit.Document;
            if (Paragraphs.Contains(action)) {
                richEdit.ContentChanged -= RichEdit_ContentChanged;
                isExternalOrder = true;
                DocumentPosition dp = document.CaretPosition;

                DocumentRange selectionRange = document.Selection;
                if (selectionRange.Length == 0)
                    selectionRange = document.Paragraphs.Get(dp).Range;

                // 선택된 범위의 시작과 끝 위치를 가져옵니다.
                int startPosition = selectionRange.Start.ToInt();
                int endPosition = selectionRange.End.ToInt();

                bool isHor = true;

                var ppAlign = ParagraphAlignment.Left;
                switch (action) {
                    case TextBarAction.HorizontalLeft:
                        ppAlign = ParagraphAlignment.Left;
                        ToggleHorizontalAlignment(action);
                        break;
                    case TextBarAction.HorizontalCenter:
                        ppAlign = ParagraphAlignment.Center;
                        ToggleHorizontalAlignment(action);
                        break;
                    case TextBarAction.HorizontalRight:
                        ppAlign = ParagraphAlignment.Right;
                        ToggleHorizontalAlignment(action);
                        break;
                    case TextBarAction.VerticalBottom:
                        isHor = false;
                        vertAct = TableCellVerticalAlignment.Bottom;
                        ToggleVerticalAlignment(action);
                        break;
                    case TextBarAction.VerticalCenter:
                        isHor = false;
                        vertAct = TableCellVerticalAlignment.Center;
                        ToggleVerticalAlignment(action);
                        break;
                    case TextBarAction.VerticalTop:
                        isHor = false;
                        vertAct = TableCellVerticalAlignment.Top;
                        ToggleVerticalAlignment(action);
                        break;
                    default: break;
                }

                var paragraphs = document.Paragraphs;
                foreach (Paragraph paragraph in paragraphs) {
                    int paragraphStart = paragraph.Range.Start.ToInt();
                    int paragraphEnd = paragraph.Range.End.ToInt();

                    if (isHor) {
                        if (IsRangeOverlap(startPosition, endPosition, paragraphStart, paragraphEnd)) {
                            paragraph.Alignment = ppAlign;
                        }
                    }

                }
                if (isClicked) RtbRtfChanged();
            }
            else if (CharactersActions.Contains(action)) {


#if false


                // dp의 텍스트와 ui상 font가 다를때 적용해야함
                // 스페이스바가 계속 끼워지는 상황

                DocumentRange selectionRange = document.Selection;
                DocumentPosition dp = document.CaretPosition;
                if (selectionRange.Length == 0 && dp.ToInt() == selectionRange.Start.ToInt()) {
                    switch (action) {
                        case TextBarAction.Name:
                            ChangeFontNameCommand fCommand = new ChangeFontNameCommand(richEdit);
                            ICommandUIState fState = fCommand.CreateDefaultCommandUIState();
                            fState.EditValue = eItem_cbxFontName.EditValue.ToString();
                            fCommand.ForceExecute(fState);
                            break;
                        case TextBarAction.Size:
                            ChangeFontSizeCommand sCommand = new ChangeFontSizeCommand(richEdit);
                            ICommandUIState sState = sCommand.CreateDefaultCommandUIState();
                            sState.EditValue = (int)eItem_cbxFontSize.EditValue;
                            sCommand.ForceExecute(sState);
                            break;
                    }

                    richEdit.Focus();
                }
                else { 

                    CharacterProperties cp = document.BeginUpdateCharacters(document.Selection);
                    bool brickColorChanged = false;
                    switch (action) {
                        case TextBarAction.Name:
                            cp.FontName = eItem_cbxFontName.EditValue.ToString();
                            break;
                        case TextBarAction.Size:
                            cp.FontSize = (int)eItem_cbxFontSize.EditValue;
                            break;
                        case TextBarAction.ForeColor:
                            cp.ForeColor = (Color)eItem_pickForeColor?.EditValue;
                            break;
                        case TextBarAction.BackColor:
                            cp.BackColor = (Color)eItem_pickBackColor.EditValue;
                            break;
                        case TextBarAction.BrickColor:
                            //brickColor = (Color)eItem_pickBrickColor.EditValue;
                            brickColorChanged = true;
                            break;
                        case TextBarAction.Bold:
                            cp.Bold = !cp.Bold ?? true;
                            break;
                        case TextBarAction.Italic:
                            cp.Italic = !cp.Italic ?? true;
                            break;
                        case TextBarAction.Underline:
                            cp.Underline = cp.Underline != UnderlineType.Single ? UnderlineType.Single : UnderlineType.None;
                            break;
                        default: break;
                    }
                    document.EndUpdateCharacters(cp);
                }
#else
                CharacterProperties cp = document.BeginUpdateCharacters(document.Selection);
                //bool brickColorChanged = false;
                switch (action) {
                    case TextBarAction.Name:
                        cp.FontName = eItem_cbxFontName.EditValue.ToString();
                        break;
                    case TextBarAction.Size:
                        cp.FontSize = (int)eItem_cbxFontSize.EditValue;
                        break;
                    case TextBarAction.ForeColor:
                        cp.ForeColor = GetFakeColor((Color)eItem_pickForeColor?.EditValue);
                        break;
                    case TextBarAction.BackColor:
                        cp.BackColor = GetFakeColor((Color)eItem_pickBackColor.EditValue);
                        break;
                    case TextBarAction.Bold:
                        cp.Bold = !cp.Bold ?? true;
                        break;
                    case TextBarAction.Italic:
                        cp.Italic = !cp.Italic ?? true;
                        break;
                    case TextBarAction.Underline:
                        cp.Underline = cp.Underline != UnderlineType.Single ? UnderlineType.Single : UnderlineType.None;
                        break;
                    default: break;
                }
                document.EndUpdateCharacters(cp);
#endif
            }
            else {
                switch (action) {
                    case TextBarAction.BrickColor:
                        var color = (Color)eItem_pickBrickColor.EditValue;
                        BrickColor = color == Color.Empty ? Color.Empty.ToArgb() : Convert.ToInt32(GetFakeColor((Color)eItem_pickBrickColor.EditValue).ToArgb());

                        if (isUserAction)
                            RtbRtfChanged();
                        else
                            isUserAction = true;
                        break;
                    default: break;
                }
            }
        }

        private Color GetFakeColor(Color color) {
            if (color.R <= 0x03 &&
                color.G <= 0x03 &&
                color.B <= 0x03) {
                color = Color.FromArgb(color.A, 3,3,3);
            }
            return color;
        }

        private void ToggleHorizontalAlignment(TextBarAction action) {
            eItem_btnHorLeft.Down = action == TextBarAction.HorizontalLeft;
            eItem_btnHorCenter.Down = action == TextBarAction.HorizontalCenter;
            eItem_btnHorRight.Down = action == TextBarAction.HorizontalRight;
        }

        private void ToggleVerticalAlignment(TextBarAction action) {
            eItem_btnVerBot.Down = action == TextBarAction.VerticalBottom;
            eItem_btnVerCenter.Down = action == TextBarAction.VerticalCenter;
            eItem_btnVerTop.Down = action == TextBarAction.VerticalTop;
        }

        private float GetUnitSize(int pixel, float dpi) {
            return DevExpress.Office.Utils.Units.PixelsToDocumentsF(pixel, dpi);
        }

        private bool IsRangeOverlap(int start1, int end1, int start2, int end2) {
            return start1 < end2 && end1 > start2;
        }

        private void RichEditSizeChanged() {
            RichControlSizeChanged = true;
        }

        private void InitializeControl() {

            
            // bar shape
            

            List<int> fontSizeArr = new List<int>();
            int value = 5;
            while (value < 500) {
                fontSizeArr.Add(value);
                if (value < 100) {
                    value++;
                }
                else if (value < 500) {
                    value += 10;
                }
            }
            repositoryItemComboBox2.Items.AddRange(fontSizeArr);
            repositoryItemComboBox2.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            eItem_cbxFontSize.EditValue = 12;

            SetRichControlMargin(richEdit, 0, 0, 0, 0);

            richEdit.PopupMenuShowing += (s, e) => RichEditContextShowing(e);
            richEdit.SizeChanged += (s, e) => RichEditSizeChanged();

            // page 외곽의 색상
            richEdit.BackColor = Color.Transparent;
            // page의 색상
            richEdit.ActiveView.BackColor = Color.FromArgb(brickColor);
            // page forecolor
            richEdit.Document.DefaultCharacterProperties.ForeColor = Color.White;
            // page backcolor
            //richEdit.Document.DefaultCharacterProperties.BackColor = Color.Transparent;


            _printRichEdit.Padding = new Padding(0);

            // page 외곽의 색상
            _printRichEdit.BackColor = Color.Transparent;
            // page의 색상
            _printRichEdit.ActiveView.BackColor = Color.Transparent;
            // page forecolor
            _printRichEdit.Document.DefaultCharacterProperties.ForeColor = Color.White;
            // page backcolor
            //_printRichEdit.Document.DefaultCharacterProperties.BackColor = Color.Transparent;
            _printRichEdit.Document.SetPageBackground(Color.Empty);

            // page 외곽의 색상
            _tmpRichEdit.BackColor = Color.Transparent;
            // page의 색상
            _tmpRichEdit.ActiveView.BackColor = Color.Transparent;
            // page forecolor
            _tmpRichEdit.Document.DefaultCharacterProperties.ForeColor = Color.White;
            // page backcolor
            //_tmpRichEdit.Document.DefaultCharacterProperties.BackColor = Color.Transparent;
            _tmpRichEdit.Document.SetPageBackground(Color.Empty);

            // 텍스트 기본 색상
            eItem_pickForeColor.EditValue = Color.White;
            var pForeColor = eItem_pickForeColor.Edit as RepositoryItemColorPickEdit;

            _tmpPrintLinkItem = new PrintableComponentLinkItem();
            _tmpPrintLinkItem.SetBackColor(Color.Empty);
            _tmpPrintSystem.ExportOptions.Image.PageBorderColor = Color.Empty;
            _tmpPrintSystem.ExportOptions.Image.PageBorderWidth = 0;
            _tmpPrintSystem.PageSettings.TopMargin = 0;
            _tmpPrintSystem.PageSettings.BottomMargin = 0;
            _tmpPrintLinkItem.SetPrintingSystem(_tmpPrintSystem);

            // font쪽 추가

            richEdit.ForeColor = Color.White;
            richEdit.Document.DefaultCharacterProperties.ForeColor = Color.White;
        }

        private void RichEditContextShowing(PopupMenuShowingEventArgs e) {
            e.Menu = null;
        }

        private void SetRichEditSize(RichEditControl richEditControl, int width, int height) {
            richEditControl.Document.Sections[0].Page.Width = DevExpress.Office.Utils.Units.PixelsToDocumentsF(width, richEditControl.DpiX);
            richEditControl.Document.Sections[0].Page.Height = DevExpress.Office.Utils.Units.PixelsToDocumentsF(height, richEditControl.DpiY);
        }
#endregion
    }

    public class TableContentIndex { 
        public int Page { get; set; }
        public int PageRow { get; set; }
        public int ParagraphPos { get; set; }
        public ParagraphAlignment PpAlignment{ get; set; }
    }

    public class TableCellInfo {
        public Color BrickColor { get; set; } = Color.Transparent;
        public float UnitWidth { get; set; }
        public float UnitHeight { get; set; }
        public TableCellVerticalAlignment CellVAlign { get; set; } = TableCellVerticalAlignment.Top;

    }

    /// <summary>
    /// RichEditControl bar버튼 옵션
    /// </summary>
    public enum TextBarAction {
        Name,
        Size,
        ForeColor,
        BackColor,
        BrickColor,
        Bold,
        Italic,
        Underline,
        HorizontalLeft,
        HorizontalCenter,
        HorizontalRight,
        VerticalBottom,
        VerticalCenter,
        VerticalTop
    }
}
