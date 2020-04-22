<?php

use Illuminate\Http\Request;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| is assigned the "api" middleware group. Enjoy building your API!
|
*/

Route::middleware('auth:api')->get('/user', function (Request $request) {
    return $request->user();
});

Route::post('/register', 'Api\AuthController@register');
Route::post('/login', 'Api\AuthController@login');

Route::group(['middleware' => 'auth:api'], function(){

    Route::post('/user-details', 'Api\AuthController@details');

    Route::post('/user-chat-get', 'Api\ChatController@get_chat');
    Route::post('/user-chat-add', 'Api\ChatController@add_message');
    Route::post('/user-chat-unread', 'Api\ChatController@get_unread_data');
    Route::post('/user-set-read', 'Api\ChatController@set_read');
    Route::post('/make-an-appointment', 'Api\AppointmentController@make_appointment');
    //Route::post('/logout', 'Api\AuthController@logout');
});



Route::post('/get-userappointment', 'Api\UserController@get_userappointment');

Route::post('/get_medicinsall', 'Api\MedialogController@medicinsallmedia');
Route::post('/get_shedule', 'Api\MedialogController@shedule');
Route::post('/get_shedulebusy', 'Api\MedialogController@shedulebusy');
Route::post('/get_servall', 'Api\MedialogController@servall');
Route::post('/import', 'ImportExcelController@importExcel');

Route::post('/additional', 'Api\AdditionalController@get_additional');

Route::post('/slider', 'Api\SliderController@slider');

Route::post( '/notify', 'NotificationController@notify');
//Route::post('/password/email', 'Api\ForgotPasswordController@sendResetLinkEmail');
//Route::post('/password/reset', 'Api\ResetPasswordController@reset');