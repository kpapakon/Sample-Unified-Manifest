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
    // Office.actions.associate('helloAction', Taskpane.run);
    // Office.actions.associate('underlinedAction', Taskpane.underline);
    // Office.context.document.addHandlerAsync(
    //   Office.EventType.DocumentSelectionChanged,
    //   Taskpane.docSelectionChangeHandler,
    // );
  }

  static async run(event: Office.AddinCommands.Event) {
    // await WordActions.sayHello('blue');
    // event.completed();
  }

  static async underline(event: Office.AddinCommands.Event) {
    // await WordActions.underlineSelection();
    // event.completed();
  }

  static async docSelectionChangeHandler(
    eventArgs: Office.DocumentSelectionChangedEventArgs,
  ) {
    const selection = await WordActions.getTextSelection(eventArgs.document);
    const isAnythingSelected = selection !== null && selection !== '';
    // Enable the button to add style only if there is a selection.
    await TaskpaneActions.enableButton('UnderlineButton', isAnythingSelected);
  }
}
