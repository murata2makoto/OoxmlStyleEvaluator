module OoxmlStyleEvaluator.OptionMonad

type OptionBuilder() =
    member __.Bind(x, f) = Option.bind f x
    member __.Return(x) = Some x
    member __.ReturnFrom(x) = x
    member __.Zero() = None
    member __.Combine(m, f) = Option.bind f m
    member __.Using(resource: 'T when 'T :> System.IDisposable, body: 'T -> Option<'U>) : Option<'U> =
        try body resource
        finally resource.Dispose()

let option = OptionBuilder()