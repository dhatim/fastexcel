## Benchmarks

All benchmarks classes must end with `Benchmark` and extend `BenchmarkLauncher`. 

    mvn clean test -Pbench
    
CSV results will be written to `target` directory with the name of the class (e.g. `ReaderBenchmark.csv`)

## Tests

Run e2e tests with

    mvn clean test -Pe2e
