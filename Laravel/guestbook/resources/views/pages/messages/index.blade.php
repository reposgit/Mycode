@extends('index')
@section('content')
<h1 class="text-center">Гостевая книга</h1>
    <hr/>
    @include('common.form')
    <hr>

        <div class="text-right"><b>Всего сообщений:</b> <i class="badge">{{$count}}</i></div>
        <br/>

        @include('pages.messages.items')
@endsection
