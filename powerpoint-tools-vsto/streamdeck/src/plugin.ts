import streamDeck from "@elgato/streamdeck";
import {
  AlignLeftAction,
  AlignCenterAction,
  AlignRightAction,
  AlignTopAction,
  AlignMiddleAction,
  AlignBottomAction,
  DistributeHorizontalAction,
  DistributeVerticalAction,
  SameWidthAction,
  SameHeightAction,
  SameSizeAction,
  RectangleAction,
  RoundedRectangleAction,
  ShadowOffAction,
  ShadowOnAction,
  BorderOffAction,
  BorderOnAction,
  MatchStyleAction,
  CleanBoxesAction,
  TestConnectionAction,
} from "./actions/powerpoint-actions.js";

streamDeck.logger.setLevel("info");

streamDeck.actions.registerAction(new AlignLeftAction());
streamDeck.actions.registerAction(new AlignCenterAction());
streamDeck.actions.registerAction(new AlignRightAction());
streamDeck.actions.registerAction(new AlignTopAction());
streamDeck.actions.registerAction(new AlignMiddleAction());
streamDeck.actions.registerAction(new AlignBottomAction());
streamDeck.actions.registerAction(new DistributeHorizontalAction());
streamDeck.actions.registerAction(new DistributeVerticalAction());
streamDeck.actions.registerAction(new SameWidthAction());
streamDeck.actions.registerAction(new SameHeightAction());
streamDeck.actions.registerAction(new SameSizeAction());
streamDeck.actions.registerAction(new RectangleAction());
streamDeck.actions.registerAction(new RoundedRectangleAction());
streamDeck.actions.registerAction(new ShadowOffAction());
streamDeck.actions.registerAction(new ShadowOnAction());
streamDeck.actions.registerAction(new BorderOffAction());
streamDeck.actions.registerAction(new BorderOnAction());
streamDeck.actions.registerAction(new MatchStyleAction());
streamDeck.actions.registerAction(new CleanBoxesAction());
streamDeck.actions.registerAction(new TestConnectionAction());

streamDeck.connect();
