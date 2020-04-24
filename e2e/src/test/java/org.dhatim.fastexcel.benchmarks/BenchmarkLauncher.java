package org.dhatim.fastexcel.benchmarks;

import org.junit.jupiter.api.Test;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.results.format.ResultFormatType;
import org.openjdk.jmh.runner.Runner;
import org.openjdk.jmh.runner.RunnerException;
import org.openjdk.jmh.runner.options.Options;
import org.openjdk.jmh.runner.options.OptionsBuilder;

import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

@BenchmarkMode(Mode.SingleShotTime)
@Warmup(iterations = 3)
@Measurement(iterations = 5)
@Fork(1)
@Threads(1)
public abstract class BenchmarkLauncher {

    @Test
    public void launchBenchmarks() throws RunnerException {
        Options options = new OptionsBuilder()
                .include(Pattern.quote(getClass().getName()))
                .shouldFailOnError(true)
                .result("target/" + getClass().getSimpleName() + ".csv")
                .resultFormat(ResultFormatType.CSV)
                .build();
        new Runner(options).run();
    }

    void readerbenchmarks() throws RunnerException {
        String foo = ReaderBenchmark.class.getName() + "\\..*";
        Options options = new OptionsBuilder().include(foo)
                .mode(Mode.SingleShotTime)
                .warmupIterations(3)
                .warmupBatchSize(1)
                .measurementIterations(5)
                .threads(1)
                .forks(1)
                .timeUnit(TimeUnit.MILLISECONDS)
                .shouldFailOnError(true)
                .resultFormat(ResultFormatType.CSV)
                .result("jmh.csv")
                .build();
        new Runner(options).run();
    }


    public void writerlaunchBenchmarks() throws Exception {
        String foo = getClass().getName() + "$";
        Options options = new OptionsBuilder().include(foo)
                .mode(Mode.SingleShotTime)
                .warmupIterations(0)
                .warmupBatchSize(1)
                .measurementIterations(1)
                .threads(1)
                .forks(0)
                .timeUnit(TimeUnit.MILLISECONDS)
                .shouldFailOnError(true)
                .resultFormat(ResultFormatType.CSV)
                .result("jmh.csv")
                .build();
        new Runner(options).run();
    }
}
