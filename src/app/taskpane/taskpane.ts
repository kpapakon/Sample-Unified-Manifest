import { Component, OnInit } from '@angular/core';
import { WordActions } from '../shared/word-actions';
import { TaskpaneActions } from '../shared/taskpane-actions';

@Component({
  selector: 'app-taskpane',
  standalone: true,
  imports: [],
  templateUrl: './taskpane.html',
  styleUrls: ['./taskpane.css'],
})
export class Taskpane implements OnInit {
  
  async ngOnInit() {
    Office.actions.associate('helloAction', Taskpane.run);
    Office.actions.associate('underlinedAction', Taskpane.underline);
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      Taskpane.docSelectionChangeHandler,
    );
    this.initMode();
    await WordActions.sayHello('orange', true);
  }

  static async run(event: Office.AddinCommands.Event) {
    await WordActions.sayHello('blue');
    event.completed();
  }

  static async underline(event: Office.AddinCommands.Event) {
    await WordActions.underlineSelection();
    event.completed();
  }

  static async docSelectionChangeHandler(
    eventArgs: Office.DocumentSelectionChangedEventArgs,
  ) {
    const selection = await WordActions.getTextSelection(eventArgs.document);
    const isAnythingSelected = selection !== null && selection !== '';
    // Enable the button to add style only if there is a selection.
    //await TaskpaneActions.enableButton('UnderlineButton', isAnythingSelected);
  }

  async onButtonClick(action: boolean) {
    // Add your logic here
    action === true ? this.setAutoOpenOn() : this.setAutoOpenOff();
    console.log('Button clicked!');
    await WordActions.sayHello(action ? 'red' : 'blue', action);
  }

  async setAutoOpenOn() {
    Office.context.document.settings.set(
      'Office.AutoShowTaskpaneWithDocument',
      true,
    );
    Office.context.document.settings.saveAsync();
  }

  setAutoOpenOff() {
    Office.context.document.settings.remove(
      'Office.AutoShowTaskpaneWithDocument',
    );
    Office.context.document.settings.saveAsync();
  }

  initMode() {
    const isReadOnly =
      Office.context.document.mode === Office.DocumentMode.ReadOnly;
    console.log('Document is in read-only mode:', isReadOnly);
  }
}
