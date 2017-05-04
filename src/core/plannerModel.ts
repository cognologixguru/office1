import { IPlannerModel } from './interfaces';
import {PlannerTasks} from '../services';

export class PlannerModel implements IPlannerModel{

    private _plannerTasks: PlannerTasks;

    constructor() {
        this._plannerTasks = new PlannerTasks();
    }

}