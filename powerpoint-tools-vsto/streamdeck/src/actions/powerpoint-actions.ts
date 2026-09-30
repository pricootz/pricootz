import { action, type KeyDownEvent, SingletonAction } from "@elgato/streamdeck";
import { sendPowerPointCommand } from "../bridge.js";

abstract class PowerPointCommandAction extends SingletonAction {
  protected abstract command: string;

  override async onKeyDown(ev: KeyDownEvent): Promise<void> {
    const ok = await sendPowerPointCommand(this.command);
    if (ok) {
      await ev.action.showOk();
    } else {
      await ev.action.showAlert();
    }
  }
}

@action({ UUID: "com.pricop.powerpoint-tools.align-left" })
export class AlignLeftAction extends PowerPointCommandAction {
  protected command = "align-left";
}

@action({ UUID: "com.pricop.powerpoint-tools.align-center" })
export class AlignCenterAction extends PowerPointCommandAction {
  protected command = "align-center";
}

@action({ UUID: "com.pricop.powerpoint-tools.align-right" })
export class AlignRightAction extends PowerPointCommandAction {
  protected command = "align-right";
}

@action({ UUID: "com.pricop.powerpoint-tools.align-top" })
export class AlignTopAction extends PowerPointCommandAction {
  protected command = "align-top";
}

@action({ UUID: "com.pricop.powerpoint-tools.align-middle" })
export class AlignMiddleAction extends PowerPointCommandAction {
  protected command = "align-middle";
}

@action({ UUID: "com.pricop.powerpoint-tools.align-bottom" })
export class AlignBottomAction extends PowerPointCommandAction {
  protected command = "align-bottom";
}

@action({ UUID: "com.pricop.powerpoint-tools.distribute-horizontal" })
export class DistributeHorizontalAction extends PowerPointCommandAction {
  protected command = "distribute-horizontal";
}

@action({ UUID: "com.pricop.powerpoint-tools.distribute-vertical" })
export class DistributeVerticalAction extends PowerPointCommandAction {
  protected command = "distribute-vertical";
}

@action({ UUID: "com.pricop.powerpoint-tools.same-width" })
export class SameWidthAction extends PowerPointCommandAction {
  protected command = "same-width";
}

@action({ UUID: "com.pricop.powerpoint-tools.same-height" })
export class SameHeightAction extends PowerPointCommandAction {
  protected command = "same-height";
}

@action({ UUID: "com.pricop.powerpoint-tools.same-size" })
export class SameSizeAction extends PowerPointCommandAction {
  protected command = "same-size";
}

@action({ UUID: "com.pricop.powerpoint-tools.rectangle" })
export class RectangleAction extends PowerPointCommandAction {
  protected command = "rectangle";
}

@action({ UUID: "com.pricop.powerpoint-tools.rounded-rectangle" })
export class RoundedRectangleAction extends PowerPointCommandAction {
  protected command = "rounded-rectangle";
}

@action({ UUID: "com.pricop.powerpoint-tools.shadow-off" })
export class ShadowOffAction extends PowerPointCommandAction {
  protected command = "shadow-off";
}

@action({ UUID: "com.pricop.powerpoint-tools.shadow-on" })
export class ShadowOnAction extends PowerPointCommandAction {
  protected command = "shadow-on";
}

@action({ UUID: "com.pricop.powerpoint-tools.border-off" })
export class BorderOffAction extends PowerPointCommandAction {
  protected command = "border-off";
}

@action({ UUID: "com.pricop.powerpoint-tools.border-on" })
export class BorderOnAction extends PowerPointCommandAction {
  protected command = "border-on";
}

@action({ UUID: "com.pricop.powerpoint-tools.match-style" })
export class MatchStyleAction extends PowerPointCommandAction {
  protected command = "match-style";
}

@action({ UUID: "com.pricop.powerpoint-tools.clean-boxes" })
export class CleanBoxesAction extends PowerPointCommandAction {
  protected command = "clean-boxes";
}

