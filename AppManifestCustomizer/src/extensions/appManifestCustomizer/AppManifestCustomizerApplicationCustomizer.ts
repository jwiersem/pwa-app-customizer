import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';

import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'AppManifestCustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppManifestCustomizerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppManifestCustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppManifestCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IAppManifestCustomizerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    
    let aLink = document.createElement("link");
    aLink.rel = "manifest";
    aLink.href = "/sites/Manifest/Shared%20Documents/manifest.webmanifest";
    aLink.crossOrigin = "use-credentials";

    document.head.insertAdjacentElement("beforeend", aLink);

    let deferredPrompt;

    // button doesn't exist yet when this runs
    const addBtn = document.querySelector('#add');    

    window.addEventListener('beforeinstallprompt', (installEvent) => {
      console.log("before install prompt raised.");

      installEvent.preventDefault();
      deferredPrompt = installEvent;      

      addBtn.addEventListener('click', (clickEvent) => {  
        console.log("Button click triggered.");
        console.log("deferredprompt" + deferredPrompt);

        deferredPrompt.prompt();

        deferredPrompt.userChoice.then((choiceResult) => {
          if (choiceResult.outcome === 'accepted') {
            console.log('User accepted the A2HS prompt');
          } else {
            console.log('User dismissed the A2HS prompt');
          }
          deferredPrompt = null;
        });
      });
    });    

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    // Handling the bottom placeholder
  if (!this._topPlaceholder) {
    this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top,
      { onDispose: this._onDispose }
    );

    // The extension should not assume that the expected placeholder is available.
    if (!this._topPlaceholder) {
      console.error("The expected placeholder (Bottom) was not found.");
      return;
    }

    if (this._topPlaceholder.domElement) {
      this._topPlaceholder.domElement.innerHTML = `
      <button id="add">Add to home</button>
      <script>if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('https://jwiersem.sharepoint.com/sites/Manifest/Shared%20Documents/sw.js').then(function(reg) {
            console.log('Successfully registered service worker', reg);
        }).catch(function(err) {
            console.warn('Error whilst registering service worker', err);
        });
      }</script>`;
    }
    
  }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }

}
