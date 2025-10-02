# Author: Cascade
# Description: Fusion 360 add-in that loads model parameter sets from a JSON file,
# applies them sequentially, and exports a STEP model for each set.

import adsk.core
import adsk.fusion
import json
import os
import traceback

_app = None
_ui = None
_handlers = []

CMD_ID = 'Cascade_ParamBatchExporter'
CMD_NAME = 'Parameter Batch Exporter'
CMD_DESCRIPTION = 'Apply parameter sets from JSON and export STEP files.'
PANEL_ID = 'SolidCreatePanel'


def run(context):
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        command_def = _ui.commandDefinitions.itemById(CMD_ID)
        if not command_def:
            command_def = _ui.commandDefinitions.addButtonDefinition(
                CMD_ID,
                CMD_NAME,
                CMD_DESCRIPTION
            )

        on_command_created = CommandCreatedEventHandler()
        command_def.commandCreated.add(on_command_created)
        _handlers.append(on_command_created)

        panel = _ui.allToolbarPanels.itemById(PANEL_ID)
        if panel is None:
            panel = _ui.allToolbarPanels.item(0)

        control = panel.controls.itemById(CMD_ID)
        if not control:
            control = panel.controls.addCommand(command_def)
        control.isPromoted = True
        control.isVisible = True
    except Exception:
        if _ui:
            _ui.messageBox('Failed to start add-in:\n{}'.format(traceback.format_exc()))


def stop(context):
    try:
        if not _app or not _ui:
            app = adsk.core.Application.get()
            ui = app.userInterface
        else:
            ui = _ui

        command_def = ui.commandDefinitions.itemById(CMD_ID)
        if command_def:
            command_def.deleteMe()

        panel = ui.allToolbarPanels.itemById(PANEL_ID)
        if panel:
            control = panel.controls.itemById(CMD_ID)
            if control:
                control.deleteMe()
    except Exception:
        if _ui:
            _ui.messageBox('Failed to stop add-in:\n{}'.format(traceback.format_exc()))


class CommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            cmd = args.command
            inputs = cmd.commandInputs

            json_path_input = inputs.addStringValueInput(
                'jsonPathInput',
                'JSON file path',
                ''
            )
            json_path_input.tooltip = 'Enter an absolute path to the parameter JSON file.'
            json_path_input.isFullWidth = True

            start_button = inputs.addBoolValueInput(
                'startButton',
                'Start',
                False,
                '',
                True
            )
            start_button.isFullWidth = True

            on_input_changed = InputChangedEventHandler(json_path_input, start_button)
            cmd.inputChanged.add(on_input_changed)
            _handlers.append(on_input_changed)

            on_destroy = DestroyHandler()
            cmd.destroy.add(on_destroy)
            _handlers.append(on_destroy)
        except Exception:
            if _ui:
                _ui.messageBox('Command creation failed:\n{}'.format(traceback.format_exc()))


class InputChangedEventHandler(adsk.core.InputChangedEventHandler):
    def __init__(self, json_path_input, start_button):
        super().__init__()
        self._json_path_input = json_path_input
        self._start_button = start_button

    def notify(self, args):
        try:
            changed_input = args.input
            if changed_input.id != self._start_button.id or not changed_input.value:
                return

            json_path = self._json_path_input.value.strip()
            if not json_path:
                _ui.messageBox('Please enter the absolute path to a JSON file.')
                changed_input.value = False
                return

            results = process_parameter_sets(json_path)

            if results.success:
                summary = '\n'.join(results.messages)
                _ui.messageBox('Parameter batch export completed successfully.\n\n{}'.format(summary))
            else:
                summary = '\n'.join(results.messages)
                _ui.messageBox('One or more exports failed.\n\n{}'.format(summary))
        except Exception:
            if _ui:
                _ui.messageBox('Processing failed:\n{}'.format(traceback.format_exc()))
        finally:
            # Reset the button so it can be pressed again without closing the command dialog.
            try:
                self._start_button.value = False
            except Exception:
                pass


class DestroyHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        # Allow the handlers list to be cleared when the command is destroyed.
        while _handlers:
            _handlers.pop()


class BatchResult:
    def __init__(self, success, messages):
        self.success = success
        self.messages = messages


def process_parameter_sets(json_path):
    app = adsk.core.Application.get()
    ui = app.userInterface

    try:
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design:
            return BatchResult(False, ['Active design not found. Open a Fusion design (.f3d/.3mf/.3d) before running the exporter.'])

        data = load_parameter_sets(json_path)
        unit = data.get('unit', 'mm')
        models = data.get('models', [])
        output_dir = data.get('outputDirectory')

        if not isinstance(models, list) or not models:
            return BatchResult(False, ['No models found in JSON file.'])

        if not output_dir:
            return BatchResult(False, ['"outputDirectory" is missing in JSON file.'])

        os.makedirs(output_dir, exist_ok=True)

        messages = []
        overall_success = True

        for index, model in enumerate(models, start=1):
            model_name = str(model.get('name') or f'Model_{index}')
            try:
                apply_parameters(design, model, unit)
                export_path = export_model(design, output_dir, model_name)
                messages.append(f'{model_name}: Exported to {export_path}')
            except Exception as process_error:
                overall_success = False
                messages.append(f'{model_name}: FAILED ({process_error})')
                ui.palettes.itemById('TextCommands').writeText(traceback.format_exc())

        return BatchResult(overall_success, messages)
    except Exception as error:
        return BatchResult(False, [f'Unexpected error: {error}\n{traceback.format_exc()}'])


def load_parameter_sets(json_path):
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f'JSON file not found: {json_path}')

    with open(json_path, 'r', encoding='utf-8') as json_file:
        return json.load(json_file)


def apply_parameters(design, model, unit):
    required_params = ['height', 'width', 'thickness']
    missing = [param for param in required_params if param not in model]
    if missing:
        raise ValueError('Missing parameters: {}'.format(', '.join(missing)))

    user_params = design.userParameters
    all_params = design.allParameters
    for name in required_params:
        value = model[name]
        expression = build_expression(value, unit)
        existing = all_params.itemByName(name)
        if existing:
            existing.expression = expression
        else:
            value_input = adsk.core.ValueInput.createByString(expression)
            user_params.add(name, value_input, unit, '')

    design.timeline.moveToEnd()


def build_expression(value, unit):
    if isinstance(value, (int, float)):
        return f'{value} {unit}'
    return str(value)


def export_model(design, output_dir, model_name):
    safe_name = sanitize_filename(model_name)
    export_path = os.path.join(output_dir, f'{safe_name}.step')
    export_mgr = design.exportManager
    options = export_mgr.createSTEPExportOptions(export_path)
    export_mgr.execute(options)
    return export_path


def sanitize_filename(name):
    invalid_chars = '<>:"/\\|?*'
    sanitized = ''.join('_' if ch in invalid_chars else ch for ch in name)
    return sanitized.strip() or 'model'


def write_text(message):
    palette = _ui.palettes.itemById('TextCommands') if _ui else None
    if palette:
        palette.writeText(message)

