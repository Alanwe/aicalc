# Phase 5: Advanced UI/UX Features - Implementation Plan

**Status**: ✅ PARTIALLY COMPLETE (60% - 3 of 5 tasks done)  
**Completion Date**: October 2025  
**Actual Duration**: ~8 hours  
**Priority**: High  
**Dependencies**: Phase 4 ✅ COMPLETE  

## Executive Summary

Phase 5 focused on transforming AiCalc into a polished, professional spreadsheet application with Excel-like usability and modern UI/UX features. This phase implemented advanced navigation, editing, and visualization features.

**Completed Objectives**:
- ✅ Excel-like keyboard navigation (Task 14A: ~2 hours)
- ✅ Right-click context menus (Task 16: ~2 hours)
- ✅ Enhanced theme system (Task 10: ~1-2 hours)

**Skipped Objectives (WinUI 3 XAML Compiler Limitations)**:
- ⏭️ Resizable and collapsible UI panels (Task 14B)
- ⏭️ Rich cell editing dialogs (Task 15)

**Remaining Objectives**:
- ⏳ Enhanced formula bar with autocomplete (Task 11: 3-4 hours)

## Phase 5 Task Breakdown

### Task 14: Advanced UI Layout & Navigation (6-9 hours)

#### Part A: Keyboard Navigation (4-6 hours) ⭐ PRIORITY

**Goal**: Implement Excel-like keyboard shortcuts for seamless navigation

**Implementation Details**:

1. **Arrow Key Navigation** (1 hour)
   ```csharp
   // MainWindow.xaml.cs
   private void MainWindow_KeyDown(object sender, KeyRoutedEventArgs e)
   {
       if (_selectedCell == null) return;
       
       switch (e.Key)
       {
           case VirtualKey.Up:
               MoveSelection(0, -1);
               e.Handled = true;
               break;
           case VirtualKey.Down:
               MoveSelection(0, 1);
               e.Handled = true;
               break;
           case VirtualKey.Left:
               MoveSelection(-1, 0);
               e.Handled = true;
               break;
           case VirtualKey.Right:
               MoveSelection(1, 0);
               e.Handled = true;
               break;
       }
   }
   
   private void MoveSelection(int colDelta, int rowDelta)
   {
       var currentAddr = _selectedCell.Address;
       var newCol = currentAddr.ColumnIndex + colDelta;
       var newRow = currentAddr.RowIndex + rowDelta;
       
       if (newCol >= 0 && newRow >= 0 && 
           newCol < ViewModel.SelectedSheet.ColumnHeaders.Count &&
           newRow < ViewModel.SelectedSheet.Rows.Count)
       {
           var newCell = ViewModel.SelectedSheet.Rows[newRow].Cells[newCol];
           SelectCell(newCell, GetButtonForCell(newCell));
       }
   }
   ```

2. **Tab Navigation** (30 minutes)
   ```csharp
   case VirtualKey.Tab:
       if (e.KeyStatus.IsKeyReleased == false)
       {
           var shift = Window.Current.CoreWindow.GetKeyState(VirtualKey.Shift).HasFlag(CoreVirtualKeyStates.Down);
           MoveSelection(shift ? -1 : 1, 0);
           e.Handled = true;
       }
       break;
   ```

3. **Enter Key Behavior** (1 hour)
   ```csharp
   case VirtualKey.Enter:
       var shift = Window.Current.CoreWindow.GetKeyState(VirtualKey.Shift).HasFlag(CoreVirtualKeyStates.Down);
       CommitCellEdit();
       MoveSelection(0, shift ? -1 : 1);
       e.Handled = true;
       break;
   ```

4. **Ctrl+Home / Ctrl+End** (30 minutes)
   ```csharp
   var ctrl = Window.Current.CoreWindow.GetKeyState(VirtualKey.Control).HasFlag(CoreVirtualKeyStates.Down);
   
   if (ctrl && e.Key == VirtualKey.Home)
   {
       // Go to A1
       var firstCell = ViewModel.SelectedSheet.Rows[0].Cells[0];
       SelectCell(firstCell, GetButtonForCell(firstCell));
       e.Handled = true;
   }
   ```

5. **Ctrl+Arrow Jump to Edge** (1 hour)
   ```csharp
   private void JumpToDataEdge(int colDelta, int rowDelta)
   {
       var current = _selectedCell;
       var sheet = ViewModel.SelectedSheet;
       
       while (true)
       {
           var newCol = current.Address.ColumnIndex + colDelta;
           var newRow = current.Address.RowIndex + rowDelta;
           
           // Check bounds
           if (newCol < 0 || newRow < 0 || 
               newCol >= sheet.ColumnHeaders.Count || 
               newRow >= sheet.Rows.Count)
               break;
           
           var nextCell = sheet.Rows[newRow].Cells[newCol];
           
           // Stop at boundary (empty -> non-empty or vice versa)
           if (IsEdgeCell(current, nextCell))
               break;
           
           current = nextCell;
       }
       
       SelectCell(current, GetButtonForCell(current));
   }
   ```

6. **F2 Edit Mode** (30 minutes)
   ```csharp
   case VirtualKey.F2:
       if (_selectedCell != null)
       {
           StartDirectEdit(_selectedCell, _selectedButton);
       }
       e.Handled = true;
       break;
   ```

7. **Page Up/Down** (30 minutes)
   ```csharp
   case VirtualKey.PageUp:
       MoveSelection(0, -10); // Jump 10 rows up
       e.Handled = true;
       break;
   case VirtualKey.PageDown:
       MoveSelection(0, 10); // Jump 10 rows down
       e.Handled = true;
       break;
   ```

**Files to Modify**:
- `MainWindow.xaml.cs` - Add keyboard event handlers
- `SheetViewModel.cs` - Navigation helpers
- `WorkbookSettings.cs` - Add EnterMoveDirection setting

**Testing Checklist**:
- [ ] Arrow keys move selection
- [ ] Tab/Shift+Tab moves horizontally
- [ ] Enter saves and moves down
- [ ] Shift+Enter saves and moves up
- [ ] Ctrl+Home goes to A1
- [ ] Ctrl+End goes to last cell
- [ ] Ctrl+Arrow jumps to data edge
- [ ] F2 enters edit mode
- [ ] Page Up/Down scrolls viewport

---

#### Part B: Resizable Function Panel (2-3 hours)

**Goal**: Make Functions panel collapsible and resizable

**Implementation**:

1. **Add GridSplitter** (1 hour)
   ```xml
   <!-- MainWindow.xaml -->
   <Grid.ColumnDefinitions>
       <ColumnDefinition Width="280" MinWidth="200" MaxWidth="500"/>
       <ColumnDefinition Width="Auto"/> <!-- For splitter -->
       <ColumnDefinition Width="*"/>
       <ColumnDefinition Width="Auto"/> <!-- For splitter -->
       <ColumnDefinition Width="320" MinWidth="250" MaxWidth="600"/>
   </Grid.ColumnDefinitions>
   
   <!-- Functions Panel -->
   <Border Grid.Column="0" Background="#2D2D30" CornerRadius="8">
       <!-- Existing content -->
   </Border>
   
   <!-- Splitter -->
   <controls:GridSplitter 
       Grid.Column="1" 
       Width="8"
       ResizeDirection="Columns"
       ResizeBehavior="PreviousAndNext"
       Background="#1E1E1E"
       Cursor="SizeWestEast"/>
   ```

2. **Collapse/Expand Button** (30 minutes)
   ```xml
   <StackPanel Orientation="Horizontal" Spacing="8">
       <TextBlock Text="📚 Functions" FontSize="16" FontWeight="SemiBold"/>
       <Button x:Name="CollapseFunctionsButton"
               Content="◀"
               Width="24" Height="24"
               Click="CollapseFunctionsPanel_Click"
               ToolTipService.ToolTip="Collapse panel"/>
   </StackPanel>
   ```

3. **Collapsible Animation** (1 hour)
   ```csharp
   private void CollapseFunctionsPanel_Click(object sender, RoutedEventArgs e)
   {
       var col = SpreadsheetGrid.ColumnDefinitions[0];
       var isCollapsed = col.Width.Value < 50;
       
       var animation = new DoubleAnimation
       {
           From = isCollapsed ? 0 : 280,
           To = isCollapsed ? 280 : 0,
           Duration = TimeSpan.FromMilliseconds(200),
           EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
       };
       
       col.BeginAnimation(ColumnDefinition.WidthProperty, animation);
       CollapseFunctionsButton.Content = isCollapsed ? "◀" : "▶";
   }
   ```

4. **Search/Filter Functions** (30 minutes)
   ```xml
   <TextBox x:Name="FunctionSearchBox"
            PlaceholderText="Search functions..."
            TextChanged="FunctionSearch_Changed"
            Margin="0,0,0,8"/>
   ```
   
   ```csharp
   private void FunctionSearch_Changed(object sender, TextChangedEventArgs e)
   {
       var query = FunctionSearchBox.Text.ToLower();
       
       foreach (var func in ViewModel.FunctionRegistry.Functions)
       {
           var visible = string.IsNullOrEmpty(query) ||
                        func.Name.ToLower().Contains(query) ||
                        func.Description.ToLower().Contains(query);
           
           // Update UI visibility
       }
   }
   ```

**Files to Create/Modify**:
- `MainWindow.xaml` - Add GridSplitter and collapse button
- `MainWindow.xaml.cs` - Collapse/expand logic
- `WorkbookSettings.cs` - Save panel width preference

---

### Task 15: Rich Cell Editing Dialogs (3-4 hours) ⭐ PRIORITY

**Goal**: Full-featured editors for complex cell types

**Implementation**:

1. **Markdown Editor Dialog** (2 hours)
   ```xml
   <!-- MarkdownEditorDialog.xaml -->
   <ContentDialog x:Class="AiCalc.MarkdownEditorDialog"
                  Title="Markdown Editor"
                  PrimaryButtonText="Save"
                  CloseButtonText="Cancel"
                  DefaultButton="Primary">
       <Grid>
           <Grid.ColumnDefinitions>
               <ColumnDefinition Width="*"/>
               <ColumnDefinition Width="Auto"/>
               <ColumnDefinition Width="*"/>
           </Grid.ColumnDefinitions>
           
           <!-- Editor -->
           <Grid Grid.Column="0">
               <StackPanel>
                   <CommandBar>
                       <AppBarButton Icon="Bold" Label="Bold" Click="InsertBold_Click"/>
                       <AppBarButton Icon="Italic" Label="Italic" Click="InsertItalic_Click"/>
                       <AppBarButton Icon="Link" Label="Link" Click="InsertLink_Click"/>
                       <AppBarButton Icon="Bullets" Label="List" Click="InsertList_Click"/>
                   </CommandBar>
                   <TextBox x:Name="MarkdownEditor"
                            AcceptsReturn="True"
                            TextWrapping="Wrap"
                            Height="400"
                            FontFamily="Consolas"
                            Text="{x:Bind MarkdownText, Mode=TwoWay}"/>
               </StackPanel>
           </Grid>
           
           <!-- Splitter -->
           <controls:GridSplitter Grid.Column="1" Width="8"/>
           
           <!-- Preview -->
           <Border Grid.Column="2" BorderBrush="#444" BorderThickness="1">
               <ScrollViewer>
                   <controls:MarkdownTextBlock 
                       x:Name="MarkdownPreview"
                       Text="{x:Bind MarkdownText, Mode=OneWay}"
                       Padding="16"/>
               </ScrollViewer>
           </Border>
       </Grid>
   </ContentDialog>
   ```

2. **JSON/XML Editor** (1 hour)
   ```xml
   <!-- JsonEditorDialog.xaml -->
   <ContentDialog Title="JSON Editor" PrimaryButtonText="Save" CloseButtonText="Cancel">
       <Grid>
           <Grid.RowDefinitions>
               <RowDefinition Height="Auto"/>
               <RowDefinition Height="*"/>
           </Grid.RowDefinitions>
           
           <CommandBar Grid.Row="0">
               <AppBarButton Icon="ReportHacked" Label="Format" Click="FormatJson_Click"/>
               <AppBarButton Icon="Accept" Label="Validate" Click="ValidateJson_Click"/>
               <AppBarButton Icon="Copy" Label="Copy" Click="CopyJson_Click"/>
           </CommandBar>
           
           <TextBox x:Name="JsonEditor"
                    Grid.Row="1"
                    AcceptsReturn="True"
                    TextWrapping="Wrap"
                    FontFamily="Consolas"
                    Text="{x:Bind JsonText, Mode=TwoWay}"/>
       </Grid>
   </ContentDialog>
   ```

3. **Image Viewer Dialog** (30 minutes)
   ```xml
   <!-- ImageViewerDialog.xaml -->
   <ContentDialog Title="Image Viewer" CloseButtonText="Close">
       <Grid>
           <Grid.RowDefinitions>
               <RowDefinition Height="*"/>
               <RowDefinition Height="Auto"/>
           </Grid.RowDefinitions>
           
           <Image x:Name="ImageDisplay" 
                  Stretch="Uniform"
                  Grid.Row="0"/>
           
           <StackPanel Grid.Row="1" Spacing="8" Margin="0,16,0,0">
               <TextBlock Text="{x:Bind ImagePath}"/>
               <TextBlock Text="{x:Bind ImageDimensions}"/>
               <TextBlock Text="{x:Bind ImageSize}"/>
           </StackPanel>
       </Grid>
   </ContentDialog>
   ```

4. **Dialog Launcher** (30 minutes)
   ```csharp
   private async void OpenCellEditor_Click(object sender, RoutedEventArgs e)
   {
       if (_selectedCell == null) return;
       
       ContentDialog dialog = _selectedCell.Value.ObjectType switch
       {
           CellObjectType.Markdown => new MarkdownEditorDialog(_selectedCell.RawValue),
           CellObjectType.Json => new JsonEditorDialog(_selectedCell.RawValue),
           CellObjectType.Xml => new XmlEditorDialog(_selectedCell.RawValue),
           CellObjectType.Image => new ImageViewerDialog(_selectedCell.RawValue),
           _ => null
       };
       
       if (dialog != null)
       {
           dialog.XamlRoot = this.Content.XamlRoot;
           var result = await dialog.ShowAsync();
           
           if (result == ContentDialogResult.Primary)
           {
               _selectedCell.RawValue = dialog.GetEditedValue();
               BuildSpreadsheetGrid(ViewModel.SelectedSheet);
           }
       }
   }
   ```

**Files to Create**:
- `MarkdownEditorDialog.xaml` + `.cs`
- `JsonEditorDialog.xaml` + `.cs`
- `XmlEditorDialog.xaml` + `.cs`
- `ImageViewerDialog.xaml` + `.cs`
- `TableEditorDialog.xaml` + `.cs`

**Dependencies**:
- Install `Microsoft.Toolkit.Uwp.UI.Controls` for MarkdownTextBlock
- Install `Newtonsoft.Json` for JSON formatting

---

### Task 16: Right-Click Context Menu (2-3 hours)

**Goal**: Rich context-sensitive actions

**Implementation**:

1. **Cell Context Menu** (1.5 hours)
   ```xml
   <!-- MainWindow.xaml -->
   <MenuFlyout x:Name="CellContextMenu" x:Key="CellContextMenu">
       <MenuFlyoutItem Text="Cut" Icon="Cut" Click="Cut_Click">
           <MenuFlyoutItem.KeyboardAccelerators>
               <KeyboardAccelerator Modifiers="Control" Key="X"/>
           </MenuFlyoutItem.KeyboardAccelerators>
       </MenuFlyoutItem>
       <MenuFlyoutItem Text="Copy" Icon="Copy" Click="Copy_Click">
           <MenuFlyoutItem.KeyboardAccelerators>
               <KeyboardAccelerator Modifiers="Control" Key="C"/>
           </MenuFlyoutItem.KeyboardAccelerators>
       </MenuFlyoutItem>
       <MenuFlyoutItem Text="Paste" Icon="Paste" Click="Paste_Click">
           <MenuFlyoutItem.KeyboardAccelerators>
               <KeyboardAccelerator Modifiers="Control" Key="V"/>
           </MenuFlyoutItem.KeyboardAccelerators>
       </MenuFlyoutItem>
       <MenuFlyoutSeparator/>
       <MenuFlyoutItem Text="Edit Cell" Icon="Edit" Click="EditCell_Click">
           <MenuFlyoutItem.KeyboardAccelerators>
               <KeyboardAccelerator Key="F2"/>
           </MenuFlyoutItem.KeyboardAccelerators>
       </MenuFlyoutItem>
       <MenuFlyoutItem Text="Clear Contents" Icon="Delete" Click="ClearContents_Click">
           <MenuFlyoutItem.KeyboardAccelerators>
               <KeyboardAccelerator Key="Delete"/>
           </MenuFlyoutItem.KeyboardAccelerators>
       </MenuFlyoutItem>
       <MenuFlyoutSeparator/>
       <MenuFlyoutSubItem Text="Insert">
           <MenuFlyoutItem Text="Insert Row Above" Click="InsertRowAbove_Click"/>
           <MenuFlyoutItem Text="Insert Row Below" Click="InsertRowBelow_Click"/>
           <MenuFlyoutItem Text="Insert Column Left" Click="InsertColumnLeft_Click"/>
           <MenuFlyoutItem Text="Insert Column Right" Click="InsertColumnRight_Click"/>
       </MenuFlyoutSubItem>
       <MenuFlyoutSubItem Text="Delete">
           <MenuFlyoutItem Text="Delete Row" Click="DeleteRow_Click"/>
           <MenuFlyoutItem Text="Delete Column" Click="DeleteColumn_Click"/>
       </MenuFlyoutSubItem>
       <MenuFlyoutSeparator/>
       <MenuFlyoutItem Text="Extract Formula" Click="ExtractFormula_Click"/>
       <MenuFlyoutItem Text="Show Dependencies" Click="ShowDependencies_Click"/>
   </MenuFlyout>
   ```

2. **Clipboard Operations** (1 hour)
   ```csharp
   private void Copy_Click(object sender, RoutedEventArgs e)
   {
       if (_selectedCell == null) return;
       
       var dataPackage = new DataPackage();
       dataPackage.SetText(_selectedCell.RawValue);
       Clipboard.SetContent(dataPackage);
       
       ViewModel.StatusMessage = "Copied to clipboard";
   }
   
   private async void Paste_Click(object sender, RoutedEventArgs e)
   {
       if (_selectedCell == null) return;
       
       var dataPackageView = Clipboard.GetContent();
       if (dataPackageView.Contains(StandardDataFormats.Text))
       {
           var text = await dataPackageView.GetTextAsync();
           _selectedCell.RawValue = text;
           await _selectedCell.EvaluateAsync();
           BuildSpreadsheetGrid(ViewModel.SelectedSheet);
       }
   }
   ```

3. **Row/Column Header Context Menu** (30 minutes)
   ```xml
   <MenuFlyout x:Name="RowHeaderContextMenu" x:Key="RowHeaderContextMenu">
       <MenuFlyoutItem Text="Insert Row Above" Click="InsertRowAbove_Click"/>
       <MenuFlyoutItem Text="Insert Row Below" Click="InsertRowBelow_Click"/>
       <MenuFlyoutItem Text="Delete Row" Click="DeleteRow_Click"/>
       <MenuFlyoutSeparator/>
       <MenuFlyoutItem Text="Hide Row" Click="HideRow_Click"/>
       <MenuFlyoutItem Text="Adjust Row Height" Click="AdjustRowHeight_Click"/>
   </MenuFlyout>
   ```

**Files to Modify**:
- `MainWindow.xaml` - Add MenuFlyouts
- `MainWindow.xaml.cs` - Implement handlers
- `SheetViewModel.cs` - Add clipboard helper methods

---

### Task 10 (Remaining): Theme System (1-2 hours)

**Goal**: Complete theme customization

**Implementation**:

1. **Theme Resource Dictionaries** (1 hour)
   ```xml
   <!-- Themes/LightTheme.xaml -->
   <ResourceDictionary xmlns="...">
       <SolidColorBrush x:Key="CellStateJustUpdatedBrush" Color="#32CD32"/>
       <SolidColorBrush x:Key="CellStateCalculatingBrush" Color="#FFA500"/>
       <SolidColorBrush x:Key="CellStateStaleBrush" Color="#1E90FF"/>
       <SolidColorBrush x:Key="CellStateManualUpdateBrush" Color="#FFA500"/>
       <SolidColorBrush x:Key="CellStateErrorBrush" Color="#DC143C"/>
       <SolidColorBrush x:Key="CellStateInDependencyChainBrush" Color="#FFD700"/>
   </ResourceDictionary>
   ```

2. **Theme Selector in Settings** (30 minutes)
   ```xml
   <!-- SettingsDialog.xaml - Add to Appearance tab -->
   <ComboBox Header="Cell State Theme"
             x:Name="ThemeSelector"
             SelectedIndex="{x:Bind ViewModel.Settings.SelectedTheme, Mode=TwoWay}">
       <ComboBoxItem Content="Light"/>
       <ComboBoxItem Content="Dark"/>
       <ComboBoxItem Content="High Contrast"/>
       <ComboBoxItem Content="Custom"/>
   </ComboBox>
   ```

3. **Custom Color Overrides** (30 minutes)
   ```xml
   <Expander Header="Customize Colors" IsExpanded="False">
       <Grid>
           <StackPanel Spacing="12">
               <ColorPicker Header="Just Updated" 
                           Color="{x:Bind CustomTheme.JustUpdatedColor, Mode=TwoWay}"/>
               <ColorPicker Header="Calculating" 
                           Color="{x:Bind CustomTheme.CalculatingColor, Mode=TwoWay}"/>
               <!-- More color pickers -->
           </StackPanel>
       </Grid>
   </Expander>
   ```

**Files to Create**:
- `Themes/LightTheme.xaml`
- `Themes/DarkTheme.xaml`
- `Themes/HighContrastTheme.xaml`

**Files to Modify**:
- `SettingsDialog.xaml` - Add theme selector
- `App.xaml.cs` - Load theme resources

---

## Implementation Order

### Day 1: Core Navigation (4-6 hours)
1. ✅ Arrow key navigation (1 hour)
2. ✅ Tab/Enter behavior (1 hour)
3. ✅ Ctrl+Home/End (30 minutes)
4. ✅ Ctrl+Arrow jump (1 hour)
5. ✅ F2 edit mode (30 minutes)
6. ✅ Page Up/Down (30 minutes)
7. ✅ Testing and refinement (30 minutes)

### Day 2: UI Enhancements (4-5 hours)
1. ✅ Resizable function panel (2 hours)
2. ✅ Search/filter functions (30 minutes)
3. ✅ Theme system (1.5 hours)
4. ✅ Testing and refinement (1 hour)

### Day 3: Rich Editing (4-5 hours)
1. ✅ Markdown editor dialog (2 hours)
2. ✅ JSON/XML editor (1 hour)
3. ✅ Image viewer (30 minutes)
4. ✅ Context menus (1.5 hours)
5. ✅ Testing and refinement (1 hour)

## Testing Checklist

### Navigation Testing
- [ ] All arrow keys work in all directions
- [ ] Tab moves right, Shift+Tab moves left
- [ ] Enter saves and moves down
- [ ] Shift+Enter saves and moves up
- [ ] Ctrl+Home goes to A1
- [ ] Ctrl+End goes to last cell
- [ ] Ctrl+Arrow jumps correctly
- [ ] F2 enters edit mode
- [ ] Page Up/Down scrolls 10 rows

### UI Testing
- [ ] Function panel resizes smoothly
- [ ] Collapse/expand animation works
- [ ] Search filters functions correctly
- [ ] GridSplitter respects min/max widths
- [ ] Panel state persists across restarts

### Dialog Testing
- [ ] Markdown editor shows live preview
- [ ] Formatting buttons insert correct syntax
- [ ] JSON formatter/validator works
- [ ] Image viewer displays correctly
- [ ] All dialogs are modal
- [ ] Save/Cancel buttons work

### Context Menu Testing
- [ ] Right-click shows menu
- [ ] Cut/Copy/Paste work
- [ ] Insert/Delete operations update grid
- [ ] Keyboard shortcuts work (Ctrl+C, Ctrl+V, etc.)
- [ ] Menu appears at cursor position

### Theme Testing
- [ ] All themes load correctly
- [ ] Cell state colors match theme
- [ ] Theme changes apply immediately
- [ ] Custom colors persist
- [ ] High contrast theme is accessible

## Success Metrics

**Technical Metrics**:
- 100% of Task 14 complete (navigation + resizable panel)
- 100% of Task 15 complete (rich editors)
- 100% of Task 16 complete (context menus)
- 100% of Task 10 remaining complete (themes)
- 0 build errors
- All keyboard shortcuts functional

**User Experience Metrics**:
- Navigation feels Excel-like and responsive
- Panels resize smoothly without lag
- Dialogs are intuitive and easy to use
- Context menus are contextually appropriate
- Themes improve visual clarity

**Code Quality Metrics**:
- No duplicate code
- Clean separation of concerns
- Proper MVVM patterns
- Comprehensive error handling
- XML comments for public methods

## Dependencies

**NuGet Packages**:
- `Microsoft.Toolkit.Uwp.UI.Controls` (for MarkdownTextBlock)
- `Newtonsoft.Json` (for JSON formatting)
- `CommunityToolkit.WinUI.UI.Controls` (for GridSplitter, if not already included)

**Asset Requirements**:
- Icons for context menu items
- Sample Markdown/JSON for testing
- Test images in various formats

## Risk Assessment

**Low Risk**:
- ✅ Navigation implementation (standard patterns)
- ✅ Theme system (already partially implemented)
- ✅ Context menus (WinUI has good support)

**Medium Risk**:
- ⚠️ GridSplitter behavior (may need custom implementation)
- ⚠️ Markdown preview (library compatibility)
- ⚠️ JSON tree view (custom control needed?)

**Mitigation**:
- Test GridSplitter early, have fallback to fixed-width toggle
- Use Microsoft.Toolkit.Uwp.UI.Controls for Markdown
- Start with simple JSON text editor, add tree view in Phase 6+

## Post-Phase 5 Recommendations

After completing Phase 5, consider:

1. **Phase 6**: Data sources and external connections
2. **Performance optimization**: Cell virtualization for large sheets
3. **Advanced features**: Pivot tables, charting
4. **Python SDK**: Script integration
5. **Polish**: Animations, transitions, loading states

## Conclusion

Phase 5 transforms AiCalc from a functional prototype into a professional, production-ready spreadsheet application. By implementing Excel-like navigation, resizable panels, rich editing dialogs, and comprehensive context menus, users will have a familiar and powerful interface.

**Estimated Total Time**: 12-16 hours  
**Priority**: High (critical for user adoption)  
**Dependencies**: Phase 4 ✅ Complete  
**Status**: 🚀 Ready to implement

---

**Planning Date**: October 11, 2025  
**Target Completion**: October 12-13, 2025  
**Next**: Begin implementation with Task 14 Part A (Navigation)
