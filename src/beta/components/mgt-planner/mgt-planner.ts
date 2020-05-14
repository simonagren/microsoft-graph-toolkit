/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import { styles } from './mgt-planner-css';

/**
 * foo
 *
 * @export
 * @class MgtPlanner
 * @extends {MgtTasksBase}
 */
@customElement('mgt-planner')
export class MgtPlanner extends MgtTasksBase {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  public static get styles() {
    return styles;
  }

  /**
   * foo
   *
   * @protected
   * @returns {import('lit-html').TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderHeaderContent(): TemplateResult {
    return html`
      header content
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {import('lit-html').TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderNewTaskDetails(): TemplateResult {
    return html`
      new task details
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {import('lit-html').TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderTasks(): TemplateResult {
    return html`
      Tasks go here
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPlanner
   */
  protected async createNewTask(): Promise<void> {
    // nope
  }
}
