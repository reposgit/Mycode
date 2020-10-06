<?php

use Illuminate\Support\Facades\Route;


/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

//Route::get('/', function () {
//    return view('welcome');
//});

Route::get('/', 'HomeController@index')->name('messages.index');
Route::post('/messagesstore','HomeController@store')->name('store');
Route::get('message/{id}/edit',['uses' => 'HomeController@edit', 'as' => 'message.edit'])->where(['id' => '[0-9]+']);
Route::get('message/{id}/destroy',['uses' => 'HomeController@destroy', 'as' => 'message.destroy'])->where(['id' => '[0-9]+']);
