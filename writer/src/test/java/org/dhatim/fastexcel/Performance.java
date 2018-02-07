/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.dhatim.fastexcel;

import java.util.concurrent.TimeUnit;
import org.junit.Test;
import org.openjdk.jmh.annotations.Mode;
import org.openjdk.jmh.results.format.ResultFormatType;
import org.openjdk.jmh.runner.Runner;
import org.openjdk.jmh.runner.options.Options;
import org.openjdk.jmh.runner.options.OptionsBuilder;

/**
 * Trigger PMH benchmarks and produce csv result file.
 */
public class Performance {

    @Test
    public void launchBenchmarks() throws Exception {
        String foo = Benchmarks.class.getName() + "\\..*";
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
