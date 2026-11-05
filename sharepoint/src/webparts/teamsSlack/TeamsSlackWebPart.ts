// Hosts teams-slack-style.html as a SharePoint full-page app. The mounting logic lives in host/mount.js
// (shared with the local harness at http://127.0.0.1:5173/sharepoint); this file only supplies the
// SharePoint pieces: the Graph token from the visitor's sign-in, SharePoint wording, and the theme.
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { AadTokenProvider } from '@microsoft/sp-http';

import { markup, runPage } from '../../generated/page';
import { installClaudeShim } from '../../generated/claudeShim';
import { mountTeamsPage } from '../../generated/mount';
import { createGraph } from '../../generated/teams-core/graph';
import { createTools } from '../../generated/teams-core/tools';
import { AuthError } from '../../generated/teams-core/errors';
import { resultError } from '../../generated/teams-core/result-error';
import { ADMIN_HINT, HELP, MESSAGES } from '../../generated/wording';

export interface ITeamsSlackWebPartProps {}

type App = ReturnType<typeof mountTeamsPage>;

export default class TeamsSlackWebPart extends BaseClientSideWebPart<ITeamsSlackWebPartProps> {
  private _app: App | undefined;
  private _dark = false;
  private _tokenProvider: Promise<AadTokenProvider> | undefined;

  protected onInit(): Promise<void> {
    this._tokenProvider = this.context.aadTokenProviderFactory.getTokenProvider();
    return Promise.resolve();
  }

  public render(): void {
    if (this._app) return; // the app manages its own state; don't rebuild on re-render
    this._app = mountTeamsPage(this.domElement, {
      page: { markup, runPage },
      installClaudeShim,
      core: { createGraph, createTools, resultError, AuthError },
      getToken: () => this._getToken(),
      help: HELP,
      messages: MESSAGES,
      dark: this._dark,
      log: (...a: unknown[]) => console.warn('[teams-slack]', ...a),
    });
  }

  /** SPFx tokens carry every Graph permission the admin approved, so the requested scopes are ignored. */
  private async _getToken(): Promise<string> {
    try {
      const provider = await this._tokenProvider!;
      return await provider.getToken('https://graph.microsoft.com');
    } catch (e) {
      const msg = String((e as Error)?.message || e);
      if (/AADSTS(65001|650057|500011)|consent|not.*approved|permission/i.test(msg)) {
        throw new AuthError('consent_missing', `SharePoint couldn’t get a Microsoft Graph token for this page. ${ADMIN_HINT} (${msg.slice(0, 160)})`);
      }
      throw new AuthError('needs_reauth', msg);
    }
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    this._dark = !!theme?.isInverted; // follow the SharePoint theme, not the OS, so the app matches the page chrome
    this._app?.setDark(this._dark);
  }

  protected onDispose(): void {
    this._app?.dispose();
    this._app = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
