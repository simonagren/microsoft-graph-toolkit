/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { ComponentMediaQuery } from '../../../components/baseComponent';
import { MgtTemplatedComponent } from '../../../components/templatedComponent';

/**
 * foo
 *
 * @export
 * @class MgtTasksBase
 * @extends {MgtTemplatedComponent}
 */
export abstract class MgtTasksBase extends MgtTemplatedComponent {
  /**
   * determines if tasks are un-editable
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean;

  /**
   * sets whether the header is rendered
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-header', type: Boolean })
  public hideHeader: boolean;

  /**
   * sets whether the options are rendered
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-options', type: Boolean })
  public hideOptions: boolean;

  /**
   * Get whether new task view is visible
   *
   * @memberof MgtTasks
   */
  public get isNewTaskVisible() {
    return this._isNewTaskVisible;
  }
  public set isNewTaskVisible(value: boolean) {
    if (value !== this._isNewTaskVisible) {
      this._isNewTaskVisible = value;
      if (!value) {
        this.clearNewTaskData();
      }
    }
  }

  /**
   * foo
   *
   * @readonly
   * @protected
   * @type {string}
   * @memberof MgtTasksBase
   */
  protected get newTaskName(): string {
    return this._newTaskName;
  }

  /**
   * foo
   *
   * @protected
   * @type {boolean}
   * @memberof MgtTasksBase
   */
  @property() protected isNewTaskBeingAdded: boolean;

  @property() private _isNewTaskVisible: boolean;
  @property() private _newTaskName: string;

  private _previousMediaQuery: ComponentMediaQuery;

  constructor() {
    super();

    this.clearState();
    this._previousMediaQuery = this.mediaQuery;
    this.onResize = this.onResize.bind(this);
  }

  /**
   * updates provider state
   *
   * @memberof MgtTasks
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
  }

  /**
   * removes updates on provider state
   *
   * @memberof MgtTasks
   */
  public disconnectedCallback() {
    window.removeEventListener('resize', this.onResize);
    super.disconnectedCallback();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const headerTemplate = !this.hideHeader ? this.renderHeader() : null;
    const newTaskTemplate = this.isNewTaskVisible ? this.renderNewTaskPanel() : null;
    const tasksTemplate = this.isLoadingState ? this.renderLoadingTask() : this.renderTasks();

    return html`
      ${headerTemplate} ${newTaskTemplate}
      <div class="Tasks">
        ${tasksTemplate}
      </div>
    `;
  }

  /**
   * Render the header part of the component.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderHeader() {
    const headerContentTemplate = this.renderHeaderContent();

    const addButton =
      !this.readOnly && !this.isNewTaskVisible
        ? html`
            <button class="AddBarItem NewTaskButton" @click="${() => (this.isNewTaskVisible = true)}">
              <span class="TaskIcon">\uE710</span>
              <span>Add</span>
            </button>
          `
        : null;

    return html`
      <div class="header">
        ${headerContentTemplate} ${addButton}
      </div>
    `;
  }

  /**
   * Render a task in a loading state.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderLoadingTask() {
    return html`
      <div class="Task LoadingTask">
        <div class="TaskContent">
          <div class="TaskCheckContainer">
            <div class="TaskCheck"></div>
          </div>
          <div class="TaskDetailsContainer">
            <div class="TaskTitle"></div>
            <div class="TaskDetails">
              <span class="TaskDetail">
                <div class="TaskDetailIcon"></div>
                <div class="TaskDetailName"></div>
              </span>
              <span class="TaskDetail">
                <div class="TaskDetailIcon"></div>
                <div class="TaskDetailName"></div>
              </span>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected renderNewTaskPanel(): TemplateResult {
    const newTaskName = this._newTaskName;

    const taskTitle = html`
      <input
        type="text"
        placeholder="Task..."
        .value="${newTaskName}"
        label="new-taskName-input"
        aria-label="new-taskName-input"
        role="input"
        @input="${(e: Event) => {
          this._newTaskName = (e.target as HTMLInputElement).value;
        }}"
      />
    `;

    const taskAddClasses = classMap({
      Disabled: !this.isNewTaskBeingAdded && (!newTaskName || !newTaskName.length),
      TaskAddButtonContainer: true
    });
    const taskAddTemplate = !this.isNewTaskBeingAdded
      ? html`
          <div class="TaskIcon TaskCancel" @click="${() => (this.isNewTaskVisible = false)}">
            <span>Cancel</span>
          </div>
          <div class="TaskIcon TaskAdd" @click="${() => this.addTask()}">
            <span>\uE710</span>
          </div>
        `
      : null;

    const newTaskDetailsTemplate = this.renderNewTaskDetails();

    return html`
      <div class="Task NewTask Incomplete">
        <div class="TaskContent">
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${taskTitle}
            </div>
            <div class="TaskDetails">
              ${newTaskDetailsTemplate}
            </div>
          </div>
        </div>
        <div class="${taskAddClasses}">
          ${taskAddTemplate}
        </div>
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderHeaderContent(): TemplateResult;

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderNewTaskDetails(): TemplateResult;

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderTasks(): TemplateResult;

  /**
   * foo
   *
   * @protected
   * @returns
   * @memberof MgtTasksBase
   */
  protected async addTask() {
    if (this.isNewTaskBeingAdded || !this.newTaskName) {
      return;
    }

    this.isNewTaskBeingAdded = true;
    await this.requestUpdate();

    try {
      await this.createNewTask();
    } finally {
      this.isNewTaskBeingAdded = false;
      this.isNewTaskVisible = false;
      this.requestUpdate();
    }
  }

  /**
   * Make a service call to create the new task object.
   *
   * @protected
   * @abstract
   * @memberof MgtTasksBase
   */
  protected abstract createNewTask(): Promise<void>;

  /**
   * Clear the form data from the new task panel.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected clearNewTaskData(): void {
    this._newTaskName = '';
  }

  /**
   * Clear the component state.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected clearState(): void {
    this.clearNewTaskData();
    this._isNewTaskVisible = false;
  }

  private onResize() {
    if (this.mediaQuery !== this._previousMediaQuery) {
      this._previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  }
}
