<?php

namespace App\Http\Controllers\Api;

use DB;
use App\User;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;

class DoctorController extends Controller
{
	public function get_doctors(Request $request)
	{
		$data = $request->all();

		if(!empty($data['category_id']))
		{
			$doctors = DB::table('doctors as d')
				->select('d.*')
				->leftJoin('doctors_services as ds', 'ds.doctor_id', '=', 'd.id')
				->where('ds.category_id', $data['category_id'])
				->groupBy('ds.doctor_id')
				->orderBy('d.order_', 'asc')
				->get();
		}
		else if(!empty($data['service_id'])&&empty($data['subservice_id']))
		{
		    $subservice_count = DB::table('doctors_services')->where('service_id', $data['service_id'])->whereNotNull('subservice_id')->count();
		    if($subservice_count != 0){
		        $doctors = [];
            }else{
			$doctors = DB::table('doctors as d')
				->select('d.*')
				->leftJoin('doctors_services as ds', 'ds.doctor_id','=', 'd.id')
				->where('ds.service_id', $data['service_id'])
				->groupBy('ds.doctor_id')
				->orderBy('d.order_', 'asc')
				->get();
		    }
		}
		else if(!empty($data['subservice_id'])&&empty($data['subsubservice_id']))
		{
            $subsubservice_count = DB::table('doctors_services')->where('subservice_id', $data['subservice_id'])->whereNotNull('subsubservice_id')->count();
            if($subsubservice_count != 0){
                $doctors = [];
            }else {
                $doctors = DB::table('doctors as d')
                    ->select('d.*')
                    ->leftJoin('doctors_services as ds', 'ds.doctor_id', '=', 'd.id')
                    ->where('ds.subservice_id', $data['subservice_id'])
                    ->orderBy('d.order_', 'asc')
                    ->get();
            }
		}
        else if(!empty($data['subsubservice_id']))
        {
                $doctors = DB::table('doctors as d')
                    ->select('d.*')
                    ->leftJoin('doctors_services as ds', 'ds.doctor_id', '=', 'd.id')
                    ->where('ds.subsubservice_id', $data['subsubservice_id'])
                    ->orderBy('d.order_', 'asc')
                    ->get();

        }
		else{
			if(!empty($data['clinic_id'])){
				$doctors = DB::table('doctors')->where('clinic_id', $data['clinic_id'])->orderBy('order_', 'asc')->get();
			}
			else{
				$doctors = DB::table('doctors')->orderBy('order_', 'asc')->get();
			}
		}
        foreach($doctors as $item) {
            if ($item->image == '') {
                $item->image = 'images/nodoctor.png';
            }
        }
		return response()->json([
			'status' => 'success',
			'data' => $doctors
		], 200);
	}

	public function get_doctor(Request $request, $id)
	{
		$doctor = DB::table('doctors')->where('id', $id)->first();
        foreach($doctor as $item) {
            if ($item->image == '') {
                $item->image = 'images/nodoctor.png';
            }
        }
		return response()->json([
			'status' => 'success',
			'data' => $doctor
		], 200);
	}

	public function get_worktime(Request $request)
	{
		$data = $request->all();

		if(!empty($data['doctor_id']))
		{
			$doctors_db = DB::table('doctors')
				->where('id', $data['doctor_id'])
				->get();
		}
		else if(!empty($data['price_id']))
		{
			$doctors_db = DB::table('doctors as d')
				->select('d.*')
				->leftJoin('doctors_services as ds', 'ds.doctor_id','=', 'd.id')
				->leftJoin('services_prices as sp', function($join){
					$join->on('sp.service_id','=','ds.service_id');
					$join->on('sp.subservice_id','=','ds.subservice_id');
				})
				->where('sp.id', $data['price_id'])
				->orderBy('d.order_', 'asc')
				->get();
		}
		else{
			return response()->json([
				'status' => 'error',
				'message' => 'invalid data'
			], 400);
		}

		$doctors_column = array_column ($doctors_db->toArray(), 'name', 'id');
		$doctors_id = array_keys($doctors_column);

		$doctors_calendar_db = DB::table('doctors_calendar')
			->whereIn('doctor_id', $doctors_id)
            ->where('clinic_id', '=', $data['clinic_id']
            )->get();

        $doctors_calendar_busy_db = DB::table('doctors_calendar_busy')
            ->whereIn('doctor_id', $doctors_id)
            ->where('clinic_id', '=', $data['clinic_id']
            )->get();

		$doctors_calendar = [];
		foreach ($doctors_calendar_db as $row)
		{
			$doctors_calendar[$row->doctor_id][$row->week_day] = [
				'start_time' => $row->start_time,
				'end_time' => $row->end_time,
				'weekend' => $row->weekend
			];
		}
        /*if(!empty($doctors_calendar))
        {return $doctors_calendar;}
        else{return 1;}*/

        $doctors_calendar_busy = [];
        foreach ($doctors_calendar_busy_db as $row)
        {
            $doctors_calendar_busy[$row->doctor_id][$row->week_day][$row->start_time] = [
                'duree' => $row->duree
            ];
        }

        $doctors_servdure_db = DB::table('doctors_servdure');

		$appointments_db = DB::table('appointment')
			->whereIn('doctor_id', $doctors_id)
            ->where('clinic_id', $data['clinic_id'])
			->whereRaw('`date` <= DATE_SUB(CURRENT_DATE, INTERVAL -14 DAY)')
			->pluck('date','doctor_id');

		$appointments = [];
		foreach ($appointments_db as $doctor_id => $date){
			$appointments[$doctor_id][] = $date;
		}

		$dates = [];
		$days_arr = ['Вс','Пн','Вт','Ср','Чт','Пт','Сб'];
		$months_arr = [
			'01' => 'января',
			'02' => 'февраля',
			'03' => 'марта',
			'04' => 'апреля',
			'05' => 'мая',
			'06' => 'июня',
			'07' => 'июля',
			'08' => 'августа',
			'09' => 'сентября',
			'10' => 'октября',
			'11' => 'ноября',
			'12' => 'декабря'
		];

		$cur_time = date('H') + 7;

		$doctors_date = [];

		foreach ($doctors_db as $doctor)
		{
			$doctor_id = $doctor->id;
            $doctor_idmedialog = DB::table('doctors')->distinct()->where('id',$doctor_id)->value('id_medialog');
			foreach(range(0,13) as $i)
			{
				$day_str = mktime(0, 0, 0, date("m"), date("d") + $i, date("Y"));
				$week_day = date('w', $day_str);
				$week_day_str = $days_arr[$week_day];
				$year = date('Y', $day_str);
				$day = date('d', $day_str);
				$month = $months_arr[date('m', $day_str)];
				$monthnum = date('m', $day_str);

				$full_date = "$week_day_str $day $month";

                $origin_date = "$year-$monthnum-$day";

				$doctor_date_list = [];
				if(!empty($doctors_calendar[$doctor_id][$origin_date])/*&&in_array($origin_date, $doctors_calendar)*/)
				{
					$doctor_day = $doctors_calendar[$doctor_id][$origin_date];
					if($doctor_day['weekend'])
					{
						$status = 'weekend';
						$list = [];
					}
					if(!empty($doctor_day['start_time']) && !empty($doctor_day['end_time']))
					{
						$start = date('H:i',strtotime($doctor_day['start_time']));
						$end = date('H:i',strtotime($doctor_day['end_time']));
						$status = 'work';
						$list = [];
						$duree = (int)DB::table('doctors_servdure')->distinct()->where([
                            ['doctor_id', '=', $doctor_idmedialog],
                            ['service_id', '=', $data['service_id']],
                            ['subservice_id', '=', $data['subservice_id']],
						    ])->value('duree');
                        $busy = false;

                        if(!empty($doctors_calendar_busy[$doctor_id][$origin_date]))
                        {
                            $busy = true;
                        }
						//for ($time = $start; $time < $end; $time = strtotime('+ '.$duree.' min'))
                        while($start < $end)
						{
                            $condition = self::check_date("$origin_date $start:00", $appointments, $doctor_id);
                            if ($condition == 'disable') {
                                $list[] = [
                                    'time' => $start . ' - ' . date('H:i',strtotime($start . ' + ' . $duree . ' min')),
                                    'condition' => 'disable'
                                ];
                            }
                            else {
                                if($origin_date == date('Y-m-d') && $cur_time >= $start)
                                {
                                    $list[] = [
                                        'time' => $start.' - '.strtotime($start.' + '.$duree.' min'),
                                        'condition' => 'disable'
                                    ];
                                }
                                else{
                                    if($busy){
                                    $doctor_day_busy = $doctors_calendar_busy[$doctor_id][$origin_date];
                                        //return $doctor_day_busy;
                                    foreach($doctor_day_busy as $key => $value) {
                                        $timebusy[date('H:i', strtotime($key))] = $value['duree'];
                                    }
                                    //return $timebusy;
                                    if(!empty($timebusy[$start])){
                                            $duree = (int)$timebusy[$start];
                                            $list[] = [
                                            'time' => $start.' - '.date('H:i',strtotime($start.' + '.$duree.' min')),
                                            'condition' => 'disable'
                                        ];
                                        //return $list;
                                    }
                                    else{
                                            $list[] = [
                                                'time' => $start . ' - ' . date('H:i',strtotime($start . ' + ' . $duree . ' min')),
                                                'condition' => 'available'
                                            ];
                                    }
                            }else{
                                        $list[] = [
                                            'time' => $start . ' - ' . date('H:i',strtotime($start . ' + ' . $duree . ' min')),
                                            'condition' => 'available'
                                        ];
                                 }
                            }
                            }
                            $start = date('H:i',strtotime($start . ' + ' . $duree . ' min'));
						}
					}
				}
				else{

					$status = 'absence';
					$list = [];
				}

				$doctors_date[$doctor_id][] = [
				    'date' => $origin_date,
					'title' => $full_date,
					'status' => $status,
					'list' => $list
				];
			}
		}

		return response()->json([
			'status' => 'success',
			'doctors' => $doctors_column,
			'doctors_date' => $doctors_date
		], 200);
	}

	public function check_date($res_date, $appointments, $doctor_id)
	{
		if(!empty($appointments[$doctor_id]) && in_array($res_date, $appointments[$doctor_id]))
		{
			return 'disable';
		}
		else{
			return 'available';
		}
	}
}