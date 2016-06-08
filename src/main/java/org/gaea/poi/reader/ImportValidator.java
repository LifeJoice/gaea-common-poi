package org.gaea.poi.reader;

import org.gaea.exception.ValidationFailedException;

import java.util.List;

/**
 * Created by iverson on 2016-6-6 11:29:32.
 */
public interface ImportValidator<T> {
    public void validate(List<T> validateData) throws ValidationFailedException;
}
