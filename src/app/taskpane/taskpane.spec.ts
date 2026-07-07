import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Taskpane } from './taskpane';

describe('Taskpane', () => {
  let component: Taskpane;
  let fixture: ComponentFixture<Taskpane>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Taskpane]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Taskpane);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
