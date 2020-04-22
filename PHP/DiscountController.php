<?php

namespace App\Http\Controllers\Api;

use DB;
use App\User;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;

class DiscountController extends Controller
{
    public function get_discount(Request $request)
	{
		$data = $request->all();

		if(empty($data['clinic_id']))
		{
			return response()->json([
				'status' => 'error',
				'message' => 'empty clinic_id'
			], 400);
		}

		$discount = DB::table('discount')
			->where('clinic_id', $data['clinic_id'])
			->where('visible', 1)
			->orderBy('order_', 'asc')
			->get();

		return response()->json([
			'status' => 'success',
			'data' => $discount
		], 200);
	}
}
