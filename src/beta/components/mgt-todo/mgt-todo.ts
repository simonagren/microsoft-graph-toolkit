/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { ComponentMediaQuery } from '../../../components/baseComponent';
import '../../../components/mgt-person/mgt-person';
import '../../../components/sub-components/mgt-arrow-options/mgt-arrow-options';
import '../../../components/sub-components/mgt-dot-options/mgt-dot-options';
import { MgtTemplatedComponent } from '../../../components/templatedComponent';
import { IGraph } from '../../../IGraph';
import { Providers } from '../../../Providers';
import { ProviderState } from '../../../providers/IProvider';
import { getShortDateString } from '../../../utils/Utils';
import { BetaGraph } from '../../BetaGraph';
import {
  createTodoTask,
  deleteTodoTask,
  getTodoTaskList,
  getTodoTaskLists,
  getTodoTasks,
  TaskStatus,
  TodoTask,
  TodoTaskList,
  updateTodoTask
} from './graph.todo';
import { styles } from './mgt-todo-css';

/**
 * Defines how a person card is shown when a user interacts with
 * a person component
 *
 * @export
 * @enum {number}
 *
 * @cssprop --tasks-header-padding - {String} Tasks header padding
 * @cssprop --tasks-header-margin - {String} Tasks header margin
 * @cssprop --tasks-title-padding - {String} Tasks title padding
 * @cssprop --tasks-plan-title-font-size - {Length} Tasks plan title font size
 * @cssprop --tasks-plan-title-padding - {String} Tasks plan title padding
 * @cssprop --tasks-new-button-width - {String} Tasks new button width
 * @cssprop --tasks-new-button-height - {String} Tasks new button height
 * @cssprop --tasks-new-button-color - {Color} Tasks new button color
 * @cssprop --tasks-new-button-background - {String} Tasks new button background
 * @cssprop --tasks-new-button-border - {String} Tasks new button border
 * @cssprop --tasks-new-button-hover-background - {Color} Tasks new button hover background
 * @cssprop --tasks-new-button-active-background - {Color} Tasks new button active background
 * @cssprop --tasks-new-task-name-margin - {String} Tasks new task name margin
 * @cssprop --task-margin - {String} Task margin
 * @cssprop --task-box-shadow - {String} Task box shadow
 * @cssprop --task-background - {Color} Task background
 * @cssprop --task-border - {String} Task border
 * @cssprop --task-header-color - {Color} Task header color
 * @cssprop --task-header-margin - {String} Task header margin
 * @cssprop --task-detail-icon-margin -{String}  Task detail icon margin
 * @cssprop --task-new-margin - {String} Task new margin
 * @cssprop --task-new-border - {String} Task new border
 * @cssprop --task-new-line-margin - {String} Task new line margin
 * @cssprop --tasks-new-line-border - {String} Tasks new line border
 * @cssprop --task-new-input-margin - {String} Task new input margin
 * @cssprop --task-new-input-padding - {String} Task new input padding
 * @cssprop --task-new-input-font-size - {Length} Task new input font size
 * @cssprop --task-new-input-active-border - {String} Task new input active border
 * @cssprop --task-new-select-border - {String} Task new select border
 * @cssprop --task-new-add-button-background - {Color} Task new add button background
 * @cssprop --task-new-add-button-disabled-background - {Color} Task new add button disabled background
 * @cssprop --task-new-cancel-button-color - {Color} Task new cancel button color
 * @cssprop --task-complete-background - {Color} Task complete background
 * @cssprop --task-complete-border - {String} Task complete border
 * @cssprop --task-complete-header-color - {Color} Task complete header color
 * @cssprop --task-complete-detail-color - {Color} Task complete detail color
 * @cssprop --task-complete-detail-icon-color - {Color} Task complete detail icon color
 * @cssprop --tasks-background-color - {Color} Task background color
 * @cssprop --task-icon-alignment - {String} Task icon alignment
 * @cssprop --task-icon-background - {Color} Task icon color
 * @cssprop --task-icon-background-completed - {Color} Task icon background color when completed
 * @cssprop --task-icon-border - {String} Task icon border styles
 * @cssprop --task-icon-border-completed - {String} Task icon border style when task is completed
 * @cssprop --task-icon-border-radius - {String} Task icon border radius
 * @cssprop --task-icon-color - {Color} Task icon color
 * @cssprop --task-icon-color-completed - {Color} Task icon color when completed
 */

/**
 * component enables the user to view, add, remove, complete, or edit tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtTasks
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-todo')
export class MgtTodo extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  public static get styles() {
    return styles;
  }

  /**
   * Get whether new task view is visible
   *
   * @memberof MgtTasks
   */
  public get isNewTaskVisible() {
    return this._isNewTaskVisible;
  }

  /**
   * Set whether new task is visible
   *
   * @memberof MgtTasks
   */
  public set isNewTaskVisible(value: boolean) {
    if (value !== this._isNewTaskVisible) {
      value ? this.showNewTaskPanel() : this.hideNewTaskPanel();
    }
  }

  /**
   * determines if tasks are un-editable
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean;

  /**
   * if set, the component will only show tasks from the target list
   * @type {string}
   */
  @property({ attribute: 'target-id', type: String })
  public targetId: string;

  /**
   * if set, the component will first show tasks from this list
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-id', type: String })
  public initialId: string;

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
   * Optional filter function when rendering tasks
   *
   * @memberof MgtTasks
   */
  public taskFilter: (task: TodoTask) => boolean;

  @property() private _isNewTaskVisible: boolean;
  @property() private _newTaskBeingAdded: boolean;
  @property() private _lists: TodoTaskList[];
  @property() private _tasks: TodoTask[];
  @property() private _hiddenTasks: string[];
  @property() private _loadingTasks: string[];
  @property() private _hasDoneInitialLoad: boolean;
  @property() private _currentList: TodoTaskList;
  @property() private _isLoadingTasks: boolean;
  @property() private _newTaskName: string;

  private _newTaskDueDate: Date;
  private _newTaskListId: string;
  private _graph: IGraph;
  private previousMediaQuery: ComponentMediaQuery;

  constructor() {
    super();
    this._graph = null;
    this._newTaskName = '';
    this._newTaskDueDate = null;
    this._newTaskListId = '';
    this._currentList = null;
    this._lists = [];
    this._tasks = [];
    this._hiddenTasks = [];
    this._loadingTasks = [];
    this._isLoadingTasks = false;
    this._hasDoneInitialLoad = false;

    this.previousMediaQuery = this.mediaQuery;
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
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtTasks
   */
  public attributeChangedCallback(name: string, oldVal: string, newVal: string) {
    super.attributeChangedCallback(name, oldVal, newVal);
    // TODO: Handle change in critical attributes
    if (name === 'data-source') {
      /*this._currentList = this.initialId;

      this._newTaskListId = '';
      this._newTaskDueDate = null;
      this._newTaskName = '';
      this._newTaskBeingAdded = false;

      this._tasks = [];
      this._lists = [];

      this._hasDoneInitialLoad = false;
      this._todoDefaultSet = false;

      this.requestStateUpdate();*/
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    let tasks = [];
    if (this._tasks) {
      tasks = this._tasks.filter(task => !this._hiddenTasks.includes(task.id));
      if (this.taskFilter) {
        tasks = tasks.filter(task => this.taskFilter(task));
      }
    }

    const headerTemplate = !this.hideHeader
      ? html`
          <div class="Header">
            ${this.renderHeader()}
          </div>
        `
      : null;

    const newTaskTemplate = this._isNewTaskVisible ? this.renderNewTaskPanel() : null;
    const loadingTemplate = this.isLoadingState || this._isLoadingTasks ? this.renderLoadingTask() : null;
    const taskTemplates = repeat(tasks, task => task.id, task => this.renderTask(task));

    return html`
      ${headerTemplate}
      <div class="Tasks">
        ${newTaskTemplate} ${loadingTemplate} ${taskTemplates}
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderNewTaskPanel() {
    const taskTitle = html`
      <input
        type="text"
        placeholder="Task..."
        .value="${this._newTaskName}"
        label="new-taskName-input"
        aria-label="new-taskName-input"
        role="input"
        @input="${(e: Event) => {
          this._newTaskName = (e.target as HTMLInputElement).value;
        }}"
      />
    `;

    const lists = this._lists.filter(
      list =>
        (this._currentList && list.id === this._currentList.id) ||
        (!this._currentList && list.id === this._newTaskListId)
    );
    if (lists.length > 0 && !this._newTaskListId) {
      this._newTaskListId = lists[0].id;
    }
    const taskList = this._currentList
      ? html`
          <span class="NewTaskBucket">
            ${this.renderBucketIcon()}
            <span>${this._currentList.displayName}</span>
          </span>
        `
      : html`
          <span class="NewTaskBucket">
            ${this.renderBucketIcon()}
            <select
              .value="${this._newTaskListId}"
              @change="${(e: Event) => {
                this._newTaskListId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${lists.map(
                list => html`
                  <option value="${list.id}">${list.displayName}</option>
                `
              )}
            </select>
          </span>
        `;

    const taskDue = null;
    /*const taskDue = html`
      <span class="NewTaskDue">
        <input
          type="date"
          label="new-taskDate-input"
          aria-label="new-taskDate-input"
          role="input"
          .value="${this.dateToInputValue(this._newTaskDueDate)}"
          @change="${(e: Event) => {
            const value = (e.target as HTMLInputElement).value;
            if (value) {
              this._newTaskDueDate = new Date(value + 'T17:00');
            } else {
              this._newTaskDueDate = null;
            }
          }}"
        />
      </span>
    `;*/

    const taskAdd = this._newTaskBeingAdded
      ? html`
          <div class="TaskAddButtonContainer"></div>
        `
      : html`
          <div class="TaskAddButtonContainer ${this._newTaskName === '' ? 'Disabled' : ''}">
            <div class="TaskIcon TaskCancel" @click="${() => this.hideNewTaskPanel()}">
              <span>Cancel</span>
            </div>
            <div class="TaskIcon TaskAdd" @click="${() => this.addTask()}">
              <span>\uE710</span>
            </div>
          </div>
        `;

    return html`
      <div class="Task NewTask Incomplete">
        <div class="TaskContent">
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${taskTitle}
            </div>
            <div class="TaskDetails">
              ${taskList} ${taskDue}
            </div>
          </div>
        </div>
        ${taskAdd}
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderHeader() {
    if (this.isLoadingState || !this._hasDoneInitialLoad) {
      return html`
        <span class="LoadingHeader"></span>
      `;
    }

    const addButton =
      !this.readOnly && !this._isNewTaskVisible
        ? html`
            <button class="AddBarItem NewTaskButton" @click="${() => this.showNewTaskPanel()}">
              <span class="TaskIcon">\uE710</span>
              <span>Add</span>
            </button>
          `
        : null;

    const list = this._lists.find(l => l.id === this.targetId);

    const listOptions = {};
    for (const l of this._lists) {
      listOptions[l.displayName] = () => this.loadTaskList(l);
    }

    const listSelect = this.targetId
      ? html`
          <span class="PlanTitle">
            ${list.displayName}
          </span>
        `
      : this._currentList
      ? html`
          <mgt-arrow-options .value="${this._currentList.displayName}" .options="${listOptions}"></mgt-arrow-options>
        `
      : null;

    return html`
      <span class="TitleCont">
        ${listSelect}
      </span>
      ${addButton}
    `;
  }

  /**
   * foo
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
   * @returns
   * @memberof MgtTodo
   */
  protected renderBucketIcon() {
    return html`
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M14 2H2V4H3H5H6H10H11H13H14V2ZM10 5H6V6H10V5ZM5 5H3V14H13V5H11V6C11 6.55228 10.5523 7 10 7H6C5.44772 7 5 6.55228 5 6V5ZM1 5H2V14V15H3H13H14V14V5H15V4V2V1H14H2H1V2V4V5Z"
          fill="#3C3C3C"
        />
      </svg>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @param {TodoTask} task
   * @returns
   * @memberof MgtTodo
   */
  protected renderTask(task: TodoTask) {
    const context = { task, list: this._currentList };

    if (this.hasTemplate('task')) {
      return this.renderTemplate('task', context, task.id);
    }

    const isCompleted = (TaskStatus as any)[task.status] === TaskStatus.completed;
    const isLoading = this._loadingTasks.includes(task.id);
    const taskCheckClasses = {
      Complete: !isLoading && isCompleted,
      Loading: isLoading,
      TaskCheck: true,
      TaskIcon: true
    };

    const taskCheckContent = isLoading
      ? html`
          \uF16A
        `
      : isCompleted
      ? html`
          \uE73E
        `
      : null;

    let taskDetailsTemplate = null;
    if (this.hasTemplate('task-details')) {
      taskDetailsTemplate = this.renderTemplate('task-details', context, `task-details-${task.id}`);
    } else {
      const taskDueTemplate = task.dueDateTime
        ? html`
            <div class="TaskDetail TaskDue">
              <span>Due ${getShortDateString(new Date(task.dueDateTime.dateTime))}</span>
            </div>
          `
        : null;

      taskDetailsTemplate = html`
        <div class="TaskTitle">
          ${task.title}
        </div>
        <div class="TaskDetail TaskBucket">
          ${this.renderBucketIcon()}
          <span>${this._currentList.displayName}</span>
        </div>
        ${taskDueTemplate}
      `;
    }

    const taskOptionsTemplate =
      !this.readOnly && !this.hideOptions
        ? html`
            <div class="TaskOptions">
              <mgt-dot-options
                .options="${{
                  'Delete Task': () => this.removeTask(task.id)
                }}"
              ></mgt-dot-options>
            </div>
          `
        : null;

    const taskClasses = classMap({
      Complete: isCompleted,
      Incomplete: !isCompleted,
      ReadOnly: this.readOnly,
      Task: true
    });
    const taskCheckContainerClasses = classMap({
      Complete: isCompleted,
      Incomplete: !isCompleted,
      TaskCheckContainer: true
    });

    return html`
      <div class=${taskClasses}>
        <div class="TaskContent" @click="${(e: Event) => this.handleTaskClick(e, task)}}">
          <span class=${taskCheckContainerClasses} @click="${(e: Event) => this.handleTaskCheckClick(e, task)}">
            <span class=${classMap(taskCheckClasses)}>
              <span class="TaskCheckContent">${taskCheckContent}</span>
            </span>
          </span>
          <div class="TaskDetailsContainer ${this.mediaQuery}">
            ${taskDetailsTemplate}
          </div>
          ${taskOptionsTemplate}
          <div class="Divider"></div>
        </div>
      </div>
    `;
  }

  /**
   * loads tasks from dataSource
   *
   * @returns
   * @memberof MgtTasks
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    const graph = provider.graph.forComponent(this);
    const betaGraph = BetaGraph.fromGraph(graph);
    this._graph = betaGraph;

    if (!this._lists || !this._lists.length) {
      const lists = this.targetId
        ? [await getTodoTaskList(this._graph, this.targetId)]
        : await getTodoTaskLists(this._graph);

      let currentList = null;
      if (lists && lists.length) {
        if (this.initialId) {
          currentList = lists.find(l => l.id === this.initialId);
        }
        if (!currentList) {
          currentList = lists[0];
        }

        this._lists = lists;
        await this.loadTaskList(currentList);
      }
    }

    this._hasDoneInitialLoad = true;
  }

  private async loadTaskList(list: TodoTaskList): Promise<void> {
    this._isLoadingTasks = true;
    this._tasks = null;
    this._currentList = list;
    this._tasks = await getTodoTasks(this._graph, list.id);
    this._isLoadingTasks = false;
  }

  private onResize() {
    if (this.mediaQuery !== this.previousMediaQuery) {
      this.previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  }

  private async updateTaskStatus(task: TodoTask, taskStatus: TaskStatus): Promise<void> {
    this._loadingTasks = [...this._loadingTasks, task.id];

    // Change the task status
    task.status = taskStatus;

    // Send update request
    const listId = this._currentList.id;
    task = await updateTodoTask(this._graph, listId, task.id, task);

    const taskIndex = this._tasks.findIndex(t => t.id === task.id);
    this._tasks[taskIndex] = task;

    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    this.requestStateUpdate();
  }

  private async removeTask(taskId: string) {
    this._hiddenTasks = [...this._hiddenTasks, taskId];

    const listId = this._currentList.id;
    await deleteTodoTask(this._graph, listId, taskId);

    this._tasks = this._tasks.filter(t => t.id !== taskId);
    this._hiddenTasks = this._hiddenTasks.filter(id => id !== taskId);
    this.requestStateUpdate();
  }

  private async addTask() {
    if (this._newTaskBeingAdded || !this._newTaskName) {
      return;
    }

    try {
      this._newTaskBeingAdded = true;

      const listId = this._currentList.id;
      const taskData = {
        title: this._newTaskName
      };

      const task = await createTodoTask(this._graph, listId, taskData);
      this._tasks.unshift(task);
    } finally {
      this._newTaskBeingAdded = false;
      this.hideNewTaskPanel();
      await this.requestStateUpdate();
    }
  }

  private handleTaskClick(e: Event, task: TodoTask) {
    this.fireCustomEvent('taskClick', { task });
    e.stopPropagation();
    e.preventDefault();
  }

  private handleTaskCheckClick(e: Event, task: TodoTask) {
    if (!this.readOnly) {
      if ((TaskStatus as any)[task.status] === TaskStatus.completed) {
        this.updateTaskStatus(task, TaskStatus.notStarted);
      } else {
        this.updateTaskStatus(task, TaskStatus.completed);
      }

      e.stopPropagation();
      e.preventDefault();
    }
  }

  private showNewTaskPanel(): void {
    this._isNewTaskVisible = true;
  }

  private hideNewTaskPanel(): void {
    this._isNewTaskVisible = false;
    this._newTaskDueDate = null;
    this._newTaskName = '';
  }

  private dateToInputValue(date: Date) {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }
}
