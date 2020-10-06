<form action= '{{ route('store') }}' method=POST id="id-form_messages">
    @csrf

    <div class="form-group">
        <label for="name">Имя: *</label>
        <input class="form-control" placeholder="Имя" name="name" type="text" id="name" value="{{$data->name}}">
    </div>

    <div class="form-group">
        <label for="email">E-mail:</label>
        <input class="form-control" placeholder="E-mail" name="email" type="email" id="email" value="{{$data->email}}">
    </div>

    <div class="form-group">
        <label for="message">Сообщение: *</label>
        <textarea class="form-control" rows="5" placeholder="Текст сообщения" name="message" cols="50"
                  id="message">{{$data->message}}</textarea>
    </div>

    <div class="form-group">
        <input class="btn btn-primary" type="submit" value="Изменить">
    </div>

</form>
