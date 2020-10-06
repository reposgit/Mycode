<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\Message;
use App\Http\Controllers\Controller;
use Validator;

class HomeController extends Controller
{
    //
    public function index(){
        $data = [
            'title' => 'Гостевая книга',
            'messages' => Message::latest()->paginate(3),
            'count' => Message::count()
        ];
        return view('pages.messages.index',$data);
    }

    public function edit($id){
        $data = Message::find($id);
        return view('pages.messages.edit',compact('data'));
    }

    public function destroy($id) {
        Message::find($id)->delete();
        return redirect()
            ->route('home')
            ->with('sessionMessage', 'Запись удалена.');
    }

    public function store(Request $request) {
        $input = $request->all();
        $validationResult = $this->validation($input);

        if (!is_null($validationResult)) {
            return $validationResult;
        } // if
        $message = new Message();
        $message->name = $input['name'];
        $message->message = $input['message'];

        if ($message->save()) {
            return redirect()
                ->route('messages.index')
                ->with('sessionMessage', 'Запись добавлена.');
        }

        abort(500);
    }
    private function validation($input, $id = NULL) {
        $validatorErrorMessages = array(
            'required' => 'Поле :attribute обязательно к заполнению',
    );

    $validator = Validator::make(
    $input,
    array(
        'name' => 'required|max:255',
        'message' => 'required',
    ),
    $validatorErrorMessages);

    if ($validator->fails()) {
        $redirectURL = ($id == NULL) ?
        route('messages.index') :
        route('messages.edit', $id);

    return redirect($redirectURL)
        ->withErrors($validator)
        ->withInput();
    } // if

return NULL;
}

}
