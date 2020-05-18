/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as GraphTypes from '@microsoft/microsoft-graph-types';
import { IGraph } from '../../../IGraph';
import { prepScopes } from '../../../utils/GraphHelpers';

// tslint:disable
export enum TaskStatus {
  notStarted,
  inProgress,
  completed,
  deferred,
  waitingOnOthers
}
// tslint:enable

/**
 * foo
 *
 * @export
 * @param {IGraph} graph
 * @param {string} planId
 * @param {string} bucketId
 * @param {string} taskId
 * @returns {Promise<void>}
 */
export async function deletePlannerTask(graph: IGraph, taskId: string): Promise<void> {
  return graph
    .api(`/planner/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .delete();
}

/**
 * foo
 *
 * @export
 * @param {IGraph} graph
 */
export async function createPlannerTask(
  graph: IGraph,
  taskData: GraphTypes.PlannerTask
): Promise<GraphTypes.PlannerTask> {
  return graph
    .api('/planner/tasks')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .post(taskData);
}

/**
 * foo
 *
 * @export
 * @param {IGraph} graph
 * @param {string} taskId
 * @param {*} peopleObj
 */
export async function assignPeopleToTask(graph: IGraph, taskId: string, people: any) {
  await updatePlannerTask(graph, taskId, { assignments: people });
}

/**
 * async promise, allows developer to set details of planner task associated with a taskId
 *
 * @param {string} taskId
 * @param {(PlannerTask)} details
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export async function updatePlannerTask(graph: IGraph, taskId: string, details: GraphTypes.PlannerTask): Promise<any> {
  return await graph
    .api(`/planner/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .patch(JSON.stringify(details));
}

/**
 * foo
 *
 * @export
 * @returns {Promise<GraphTypes.PlannerPlan[]>}
 */
export async function getPlannerPlan(graph: IGraph, planId: string): Promise<GraphTypes.PlannerPlan> {
  const plan = await graph
    .api(`/planner/plans/${planId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return plan;
}

/**
 * foo
 *
 * @export
 * @returns {Promise<GraphTypes.PlannerPlan[]>}
 */
export async function getMyPlannerPlans(graph: IGraph): Promise<GraphTypes.PlannerPlan[]> {
  const plans = await graph
    .api('/me/planner/plans')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return plans && plans.value;
}

/**
 * foo
 *
 * @export
 * @returns {Promise<GraphTypes.PlannerBucket[]>}
 */
export async function getPlannerBucket(
  graph: IGraph,
  planId: string,
  bucketId: string
): Promise<GraphTypes.PlannerBucket> {
  const buckets = await graph
    .api(`/planner/plans/${planId}/buckets/${bucketId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return buckets && buckets.value;
}

/**
 * foo
 *
 * @export
 * @returns {Promise<GraphTypes.PlannerBucket[]>}
 */
export async function getPlannerBuckets(graph: IGraph, planId: string): Promise<GraphTypes.PlannerBucket[]> {
  const buckets = await graph
    .api(`/planner/plans/${planId}/buckets`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return buckets && buckets.value;
}

/**
 * foo
 *
 * @export
 * @returns {Promise<GraphTypes.PlannerTask[]>}
 */
export async function getPlannerTasks(graph: IGraph, bucketId: string): Promise<GraphTypes.PlannerTask[]> {
  const tasks = await graph
    .api(`/planner/buckets/${bucketId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return tasks && tasks.value;
}
