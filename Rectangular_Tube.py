# -*- coding: utf-8 -*-
#Author-Notta
#Description-
import adsk.core, adsk.fusion, adsk.cam, traceback

# グローバル変数の初期化(宣言)
_app = adsk.core.Application.cast(None)
_ui = adsk.core.UserInterface.cast(None)

#コマンドの入力用変数
_width = adsk.core.ValueCommandInput.cast(None)
_height = adsk.core.ValueCommandInput.cast(None)
_length = adsk.core.ValueCommandInput.cast(None)
_thickness = adsk.core.ValueCommandInput.cast(None)
_errMessage = adsk.core.TextBoxCommandInput.cast(None)

_handlers = []

# アドイン起動時の処理
def run(context):
    try:
        global _app, _ui
        _app = adsk.core.Application.get()
        _ui = _app.userInterface
        
        # コマンドの作成
        cmdDef = _ui.commandDefinitions.addButtonDefinition('Rectangular_Tube_PythonAddIn','中空角棒作成','中空角棒を新しいコンポーネントで作る。', 'resources/Rectangular_Tube')
        createPanel = _ui.allToolbarPanels.itemById('SolidCreatePanel')
        createPanel.controls.addCommand(cmdDef)
        
        # コマンド作成時のイベントハンドラを紐付け
        onCommandCreated = RTCommandCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        _handlers.append(onCommandCreated)
        
        # 初回起動時
        if context['IsApplicationStartup'] == False:
            _ui.messageBox('「中空角棒作成」コマンドが作成されました。\nモデルワークスペースの作成にあります。','Rectamgular Tube')

    except:
          if _ui:
              _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# アドイン終了時の処理
def stop(context):
    try:
        createPanel = _ui.allToolbarPanels.itemById('SolidCreatePanel')
        RTButton = createPanel.controls.itemById('Rectangular_Tube_PythonAddIn')       
        if RTButton:
            RTButton.deleteMe()
        
        cmdDef = _ui.commandDefinitions.itemById('Rectangular_Tube_PythonAddIn')
        if cmdDef:
            cmdDef.deleteMe()
    except:
          if _ui:
              _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))       


# コマンド作成時のイベントハンドラ
class RTCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            
            des = adsk.fusion.Design.cast(_app.activeProduct)
            if not des:
                # モデルの作業スペースでコマンドが実行されなかった場合
                _ui.messageBox('It is not supported in current workspace.\nPlease change to MODEL workspace and try again.')
                return()
            
            # 各アトリビュートの作成(groupName:RectangularTube)
            # 各デフォルト値の単位はcm
            width = '1'            
            widthAttrib = des.attributes.itemByName('RectangularTube', 'width')
            if widthAttrib:
                # 前回コマンド実行時の値がある時は、それを使用
                width = widthAttrib.value
            
            height = '1'            
            heightAttrib = des.attributes.itemByName('RectangularTube', 'height')
            if heightAttrib:
                height = heightAttrib.value
            
            length = '10'            
            lengthAttrib = des.attributes.itemByName('RectangularTube', 'length')
            if lengthAttrib:
                length = lengthAttrib.value
            
            thickness = '0.1'            
            thicknessAttrib = des.attributes.itemByName('RectangularTube', 'thickness')
            if thicknessAttrib:
                thickness = thicknessAttrib.value
            
            cmd = eventArgs.command
            # ユーザが先に別のコマンドを実行した場合の処理(True:実行,False:終了)
            cmd.isExecutedWhenPreEmpted = False
            inputs = cmd.commandInputs
            
            global _width, _height, _length, _thickness, _errMessage
            
            # コマンドダイアログの定義
            _width = inputs.addValueInput('width', '幅', 'mm', adsk.core.ValueInput.createByReal(float(width)))
            _height = inputs.addValueInput('height', '高さ', 'mm', adsk.core.ValueInput.createByReal(float(height)))
            _length = inputs.addValueInput('length', '長さ', 'mm', adsk.core.ValueInput.createByReal(float(length)))
            _thickness = inputs.addValueInput('thickness', '厚み', 'mm', adsk.core.ValueInput.createByReal(float(thickness)))
            _errMessage = inputs.addTextBoxCommandInput('errMessage', '', '', 2, True)
            _errMessage.isFullWidth = True
            
            # コマンド関係の各イベントを紐付け
            onExecute = RTCommandExecuteHandler()
            cmd.execute.add(onExecute)
            _handlers.append(onExecute)        
            
            onInputChanged = RTCommandInputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            _handlers.append(onInputChanged)     
            
            onValidateInputs = RTCommandValidateInputsHandler()
            cmd.validateInputs.add(onValidateInputs)
            _handlers.append(onValidateInputs)

        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# コマンド実行時のイベントハンドラ
class RTCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # 現在の値をアトリビュートとして保存
            des = adsk.fusion.Design.cast(_app.activeProduct)
            attribs = des.attributes
            attribs.add('RectangularTube', 'width', str(_width.value))
            attribs.add('RectangularTube', 'height', str(_height.value))
            attribs.add('RectangularTube', 'length', str(_length.value))
            attribs.add('RectangularTube', 'thickness', str(_thickness.value))
            
            # 現在の値を取得
            width = _width.value
            height = _height.value
            length = _length.value
            thickness = _thickness.value
            
            # 中空角棒(コンポーネント)の作成
            RTComp = makeRectangularTube(des, width, height, length, thickness)
            
            # コンポーネントの説明変更
            desc = str(width) + "×" + str(height) + "×" + str(length)
            desc += "×t" + str(thickness) + "の中空角棒"
            RTComp.description = desc
            
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))            


# コマンドの入力変化時のイベントハンドラ
class RTCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # 特に何もしない。(今後追加予定)
            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))            


# コマンドの入力が正しいか確かめる時のイベントハンドラ
class RTCommandValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
            _errMessage.text = ''

            width = _width.value
            height = _height.value
            length = _length.value
            thickness = _thickness.value
            
            # 厚さが幅か高さの小さい方の半分より大きい時
            if width <= height:
                if width/2 < thickness:
                    _errMessage.text = '厚さが大きすぎます。幅の半分以下にしてください。'
                    eventArgs.areInputsValid = False
                    return
            else:
                if height/2 < thickness:
                    _errMessage.text = '厚さが大きすぎます。高さの半分以下にしてください。'
                    eventArgs.areInputsValid = False
                    return
            
            # どれかの値が0の時
            if width <= 0:
                _errMessage.text = '幅が0以下になっています。修正してください。'
                eventArgs.areInputsValid = False
                return
            if height <= 0:
                _errMessage.text = '高さが0以下になっています。修正してください。'
                eventArgs.areInputsValid = False
                return
            if length <= 0:
                _errMessage.text = '長さが0以下になっています。修正してください。'
                eventArgs.areInputsValid = False
                return
            if thickness <= 0:
                _errMessage.text = '厚さが0以下になっています。修正してください。'
                eventArgs.areInputsValid = False
                return
            
        except:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))            


# 中空角棒を新しいコンポーネントで作成する関数
def makeRectangularTube(design, width, height, length, thickness):
    try:
        # Create a new component by creating an occurrence.
        occs = design.rootComponent.occurrences
        mat = adsk.core.Matrix3D.create()
        newOcc = occs.addNewComponent(mat)        
        newComp = adsk.fusion.Component.cast(newOcc.component)
        
        # 新規スケッチの作成
        sketches = newComp.sketches
        xyPlane = newComp.xYConstructionPlane
        sketch = sketches.add(xyPlane)
        
        # 基準の四角形を描く。
        lines = sketch.sketchCurves.sketchLines
        rectLines = lines.addCenterPointRectangle(adsk.core.Point3D.create(0,0,0), adsk.core.Point3D.create(width/2,height/2,0))

        # Add the geometry to a collection. This uses a utility function that automatically finds the connected curves and returns a collection.
        curves = sketch.findConnectedCurves(rectLines.item(0))
               
        # オフセットを作成。
        dirPoint = adsk.core.Point3D.create(0, 0, 0)
        sketch.offset(curves, dirPoint, thickness)
        
        #プロファイルを取得
        prof = sketch.profiles.item(0)
        
        
        extrudes = newComp.features.extrudeFeatures
        #押し出し入力(引数)を作成
        extInput = extrudes.createInput(prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
        distance = adsk.core.ValueInput.createByReal(length)
        #extInput.setDistanceExtent(False, distance) # 片側だけ押し出す
        extInput.setSymmetricExtent(distance,True) # 左右対称に押し出す
        extInput.isSolid = True
        # 押し出し
        extrudes.add(extInput)
        
        # 名前変更
        newComp.name = str(width*10) + '×' + str(height*10) + '×' + str(length*10) + '中空角棒'
        return newComp
        
    except Exception as error:
        _ui.messageBox("makeRectangularTube Failed : " + str(error)) 
        return None            
