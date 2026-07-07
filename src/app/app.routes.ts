import { Routes } from '@angular/router';
import { Taskpane } from './taskpane/taskpane';

export const routes: Routes = [
  { path: 'taskpane', component: Taskpane },
  { path: '', redirectTo: 'taskpane', pathMatch: 'full' }
];
