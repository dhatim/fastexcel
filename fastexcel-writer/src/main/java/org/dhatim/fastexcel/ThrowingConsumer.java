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

import java.io.IOException;

/**
 * Consumer throwing {@link IOException}.
 *
 * @param <T> Type being consumed.
 */
@FunctionalInterface
interface ThrowingConsumer<T> {

    /**
     * Consume object.
     *
     * @param t Object being consumed.
     * @throws IOException If an I/O error occurs.
     */
    void accept(T t) throws IOException;
}
