/*
AutoHotkey

Copyright 2003-2006 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

//////////////////////////////////////////////////////////////////////////////////////
// v1.0.40.02: This is now a separate file to allow its compiler optimization settings
// to be set independently of those of the other modules.  In one benchmark, this
// improved performance of expressions and function calls by 9% (that is, when the
// other modules are set to "minmize size" such as for the AutoHotkeySC.bin file).
// This gain in performance is at the cost of a 1.5 KB increase in the size of the
// compressed code, which seems well worth it given how often expressions and
// function-calls are used (such as in loops).
//
// ExpandArgs() and related functions were also put into this file because that
// further improves performance across the board -- even for AutoHotkey.exe despite
// the fact that the only thing that changed for it was the module move, not the
// compiler settings.  Apparently, the butterfly effect can cause even minor
// modifications to impact the overall performance of the generated code by as much as
// 7%.  However, this might have more to do with cache hits and misses in the CPU than
// with the nature of the code produced by the compiler.
// UPDATE 10/18/2006: There's not much difference anymore -- in fact, using min size
// for everything makes compiled scripts slightly faster in basic benchmarks, probably
// due to the recent addition of the linker optimization that physically orders
// functions in a better order inside the EXE.  Therefore, script_expression.cpp no
// longer has a separate "favor fast code" option.
//////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h" // pre-compiled headers
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "qmath.h" // For ExpandExpression()

// __forceinline: Decided against it for this function because alhough it's only called by one caller,
// testing shows that it wastes stack space (room for its automatic variables would be unconditionally 
// reserved in the stack of its caller).  Also, the performance benefit of inlining this is too slight.
// Here's a simple way to verify wasted stack space in a caller that calls an inlined function:
//    DWORD stack
//    _asm mov stack, esp
//    MsgBox(stack);
char *Line::ExpandExpression(int aArgIndex, ResultType &aResult, char *&aTarget, char *&aDerefBuf
	, size_t &aDerefBufSize, char *aArgDeref[], size_t aExtraSize)
// Caller should ignore aResult unless this function returns NULL.
// Returns a pointer to this expression's result, which can be one of the following:
// 1) NULL, in which case aResult will be either FAIL or EARLY_EXIT to indicate the means by which the current
//    quasi-thread was terminated as a result of a function call.
// 2) The constant empty string (""), in which case we do not alter aTarget for our caller.
// 3) Some persistent location not in aDerefBuf, namely the mContents of a variable or a literal string/number,
//    such as a function-call that returns "abc", 123, or a variable.
// 4) At the offset aTarget minus aDerefBuf inside aDerefBuf (note that aDerefBuf might have been reallocated by us).
// aTarget is left unchnaged except in case #4, in which case aTarget has been adjusted to the position after our
// result-string's terminator.  In addition, in case #4, aDerefBuf, aDerefBufSize, and aArgDeref[] have been adjusted
// for our caller if aDerefBuf was too small and needed to be enlarged.
// Thanks to Joost Mulders for providing the expression evaluation code upon which this function is based.
{
	// This is the location in aDerefBuf the caller told us is ours.  Caller has already ensured that
	// our part of the buffer is large enough for our first stage expansion, but not necessarily
	// for our final result (if too large, we will expand the buffer to handle the result).
	char *target = aTarget;

	// The following must be defined early so that mem_count is initialized and the array is guaranteed to be
	// "in scope" in case of early "goto end/fail".
	#define MAX_EXPR_MEM_ITEMS 100 // Hard to imagine using even a few in a typical script, let alone 100.
	char *mem[MAX_EXPR_MEM_ITEMS]; // No init necessary.  In most cases, it will never be used.
	int mem_count = 0; // The actual number of items in use in the above array.
	char *result_to_return = ""; // By contrast, NULL is used to tell the caller to abort the current thread.  That isn't done for normal syntax errors, just critical conditions such as out-of-memory.
	Var *output_var = (mActionType == ACT_ASSIGNEXPR) ? OUTPUT_VAR : NULL; // Resolve early because it's similar in usage/scope to the above.  Plus MUST be resolved prior to calling any script-functions since they could change the values in sArgVar[].

	map_item map[MAX_DEREFS_PER_ARG*2 + 1];
	int map_count = 0;
	// Above sizes the map to "times 2 plus 1" to handle worst case, which is -y + 1 (raw+deref+raw).
	// Thus, if this particular arg has the maximum number of derefs, the number of map markers
	// needed would be twice that, plus one for the last raw text's marker.

	///////////////////////////////////////////////////////////////////////////////////////
	// EXPAND DEREFS and make a map that indicates the positions in the buffer where derefs
	// vs. raw text begin and end.
	///////////////////////////////////////////////////////////////////////////////////////
	char *pText, *this_marker;
	DerefType *deref;
	for (pText = mArg[aArgIndex].text  // Start at the begining of this arg's text.
		, deref = mArg[aArgIndex].deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG:
		DerefType &this_deref = *deref; // For performance.
		if (pText < this_deref.marker)
		{
			map[map_count].type = EXP_RAW;
			map[map_count].marker = target;  // Indicate its position in the buffer.
			// Copy the chars that occur prior to this_deref.marker into the buffer:
			for (this_marker = this_deref.marker; pText < this_marker; *target++ = *pText++); // this_marker is used to help performance.
			map[map_count].end = target; // Since RAWS are never empty due to the check above, this will always be the character after the last.
			++map_count;
		}

		// Known issue: If something like %A_Space%String exists in the script (or any variable containing
		// spaces), the expression will yield inconsistent results.  Since I haven't found an easy way
		// to fix that, not fixing it seems okay in this case because it's not a correct way to be using
		// dynamically built variable names in the first place.  In case this will be fixed in the future,
		// either directly or as a side-effect of other changes, here is a test script that illustrates
		// the inconsistency:
		//vText = ABC 
		//vNum = 1 
		//result1 := (vText = %A_space%ABC) AND (vNum = 1)
		//result2 := vText = %A_space%ABC AND vNum = 1
		//MsgBox %result1%`n%result2%

		map_item &this_map_item = map[map_count]; // For performance, but be careful not to use after ++map_count, etc.
		if (this_deref.is_function)
		{
			this_map_item.type = EXP_DEREF_FUNC;
			this_map_item.deref = deref;
			// But nothing goes into target, so this is an invisible item of sorts.
			// However, everything after the function's name, starting at its open-paren, will soon be
			// put in as a collection of normal items (raw text and derefs).
		}
		else
		{
			// GetExpandedArgSize() relies on the fact that we only expand the following items into
			// the deref buffer:
			// 1) Derefs whose var type isn't VAR_NORMAL or who have zero length (since they might be env. vars).
			// 2) Derefs that are enclosed by the g_DerefChar character (%), which in expressions means that
			//    must be copied into the buffer to support double references such as Array%i%.
			// Now copy the contents of the dereferenced var.  For all cases, the target buf has already
			// been verified to be large enough, assuming the value hasn't changed between the time we
			// were called and the time the caller calculated the space needed.
			if (*this_deref.marker == g_DerefChar)
				this_map_item.type = EXP_DEREF_DOUBLE;
			else // SINGLE or VAR.  Set initial guess to possibly be overridden later:
				this_map_item.type = (this_deref.var->Type() == VAR_NORMAL) ? EXP_DEREF_VAR : EXP_DEREF_SINGLE;

			if (this_map_item.type == EXP_DEREF_VAR)
			{
				// Need to distinguish between empty variables and environment variables because the former
				// we want to pass by reference into functions but the latter need to go into the deref buffer.
				// So if this deref's variable is of zero length: if Get() actually retrieves anything, it's
				// an environment variable rather than a zero-length normal variable. The size estimator knew
				// that and already provided space for it in the buffer.  But if it returns an empty string,
				// it's a normal empty variable and thus it stays of type EXP_DEREF_VAR.
				if (g_NoEnv || this_deref.var->Length()) // v1.0.43.08: Added g_NoEnv.
					this_map_item.var = this_deref.var;
				else // Auto-env retrieval is in effect and this var is zero-length, so check if it's an environment variable.
				{
					this_map_item.marker = target;  // Indicate its position in the buffer.
					target += this_deref.var->Get(target);
					if (this_map_item.marker == target) // Empty string, so it's not an environment variable.
						this_map_item.var = this_deref.var;
					else // Override it's original EXP_DEREF_VAR type.
					{
						this_map_item.end = target;
						this_map_item.type = EXP_DEREF_SINGLE;
					}
				}
			}
			else // SINGLE or DOUBLE, both of which need to go into the buffer.
			{
				this_map_item.marker = target;  // Indicate its position in the buffer.
				target += this_deref.var->Get(target);
				this_map_item.end = target;
				// For performance reasons, the expression parser relies on an extra space to the right of each
				// single deref.  For example, (x=str), which is seen as (x_contents=str_contents) during
				// evaluation, would instead me seen as (x_contents =str_contents ), which allows string
				// terminators to be put in place of those two spaces in case either or both contents-items
				// are strings rather than numbers (such termination also simplifies number recognition).
				// GetExpandedArgSize() has already ensured there is enough room in the deref buffer for these:
			}
			// Fix for v1.0.35.04: Each EXP_DEREF_VAR now gets a corresponding empty string in the buffer
			// as a placeholder, which prevents an expression such as x*y*z from being seen as having
			// two adjacent asterisks, which prevents it from being seen as SYM_POWER and other mistakes.
			// This could have also been solved by having SYM_POWER and other double-symbol operators
			// check to ensure the second symbol isn't at or beyond map[].end, but that would complicate
			// the code and decrease maintainability, so this method seems better.  Also note that this
			// fix isn't needed for EXP_DEREF_FUNC because the functions parentheses and arg list are
			// always present in the deref buffer, which prevents SYM_POWER and similar from seeing
			// the character after the first operator symbol as something that changes the operator.
			if (this_map_item.type != EXP_DEREF_DOUBLE) // EXP_DEREF_VAR or EXP_DEREF_SINGLE.
				*target++ = '\0'; // Always terminated since they can't form a part of a double-deref.
			// For EXP_DEREF_VAR, if our caller will be assigning the result of our expression to
			// one of the variables involved in the expression, that should be okay because:
			// 1) The expression's result is normally not EXP_DEREF_VAR because any kind of operation
			//    that is performed, such as addition or concatenation, would have transformed it into
			//    SYM_OPERAND, SYM_STRING, SYM_INTEGER, or SYM_FLOAT.
			// 2) If the result of the expression is the exact same address as the contents of the
			//    variable our caller is assigning to (which can happen from something like
			//    GlobalVar := YieldGlobalVar()), Var::Assign() handles that by checking if they're
			//    the same and also using memmove(), at least when source and target overlap.
		} // Not a function.
		++map_count; // i.e. don't increment until after we're done using the old value.
		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
		pText += this_deref.length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	if (*pText)
	{
		map[map_count].type = EXP_RAW;
		map[map_count].marker = target;  // Indicate its position in the buffer.
		for (; *pText; *target++ = *pText++);
		map[map_count].end = target;
		++map_count;
	}

	// Terminate the buffer, even if nothing was written into it:
	*target++ = '\0'; // Target must be incremented to point to the next available position (if any) for use further below.
	// The following is conservative because the original size estimate for our portion might have
	// been inflated due to:
	// 1) Falling back to MAX_FORMATTED_NUMBER_LENGTH as the estimate because the other was smaller.
	// 2) Some of the derefs being smaller than their estimate (which is a documented possibility for some built-in variables).
	size_t capacity_of_our_buf_portion = target - aTarget + aExtraSize; // The initial amount of size available to write our final result.

/////////////////////////////////////////

	// Having a precedence array is required at least for SYM_POWER (since the order of evaluation
	// of something like 2**1**2 does matter).  It also helps performance by avoiding unnecessary pushing
	// and popping of operators to the stack. This array must be kept in sync with "enum SymbolType".
	// Also, dimensioning explicitly by SYM_COUNT helps enforce that at compile-time:
	static int sPrecedence[SYM_COUNT] =
	{
		0, 0, 0, 0, 0, 0 // SYM_STRING, SYM_INTEGER, SYM_FLOAT, SYM_VAR, SYM_OPERAND, SYM_BEGIN (SYM_BEGIN must be lowest precedence).
		, 1, 1, 1        // SYM_CPAREN, SYM_OPAREN, SYM_COMMA (to simplify the code, parentheses must be lower than all operators in precedence).
		, 2              // SYM_OR
		, 3              // SYM_AND
		, 4              // SYM_LOWNOT (the low precedence version of logical-not)
		, 5, 5, 5        // SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL (lower prec. than the below so that "x < 5 = var" means "result of comparison is the boolean value in var".
		, 6, 6, 6, 6     // SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE
		, 7              // SYM_CONCAT
		, 8              // SYM_BITOR -- Seems more intuitive to have these three higher in prec. than the above, unlike C and Perl, but like Python.
		, 9              // SYM_BITXOR
		, 10             // SYM_BITAND
		, 11, 11         // SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT
		, 12, 12         // SYM_PLUS, SYM_MINUS
		, 13, 13, 13     // SYM_TIMES, SYM_DIVIDE, SYM_FLOORDIVIDE
		, 14, 14, 14, 14 // SYM_NEGATIVE (unary minus), SYM_HIGHNOT (the high precedence "not" operator), SYM_BITNOT, SYM_ADDRESS
		, 15             // SYM_POWER (see note below).
		, 16             // SYM_DEREF -- Giving this a higher precedence than the above allows !*Var to work, and also -*Var and ~*Var.
		, 17             // SYM_FUNC -- Probably must be of highest precedence for it to work properly.
	};
	// Most programming languages give exponentiation a higher precedence than unary minus and logical-not.
	// For example, -2**2 is evaluated as -(2**2), not (-2)**2 (the latter is unsupported by qmathPow anyway).
	// However, this rule requires a small workaround in the postfix-builder to allow 2**-2 to be
	// evaluated as 2**(-2) rather than being seen as an error.  v1.0.45: A similar thing is required
	// to allow the following to work: 2**!1, 2**not 0, 2**~0xFFFFFFFE, 2**&x.
	// On a related note, the right-to-left tradition of something like 2**3**4 is not implemented.
	// Instead, the expression is evaluated from left-to-right (like other operators) to simplify the code.

	#define MAX_TOKENS 512 // Max number of operators/operands.  Seems enough to handle anything realistic, while conserving call-stack space.
	ExprTokenType infix[MAX_TOKENS], *postfix[MAX_TOKENS], *stack[MAX_TOKENS + 1];  // +1 for SYM_BEGIN on the stack.
	int infix_count = 0, postfix_count = 0, stack_count = 0;
	// Above dimensions the stack to be as large as the infix/postfix arrays to cover worst-case
	// scenarios and avoid having to check for overflow.  For the infix-to-postfix conversion, the
	// stack must be large enough to hold a malformed expression consisting entirely of operators
	// (though other checks might prevent this).  It must also be large enough for use by the final
	// expression evaluation phase, the worst case of which is unknown but certainly not larger
	// than MAX_TOKENS.

	///////////////////////////////////////////////////////////////////////////////////////////////
	// TOKENIZE THE INFIX EXPRESSION INTO AN INFIX ARRAY: Avoids the performance overhead of having
	// to re-detect whether each symbol is an operand vs. operator at multiple stages.
	///////////////////////////////////////////////////////////////////////////////////////////////
	SymbolType sym_prev;
	char *op_end, *cp, *terminate_string_here;
	UINT op_length;
	Var *found_var;
	size_t literal_string_length;

	for (int map_index = 0; map_index < map_count; ++map_index) // For each deref and raw item in map.
	{
		// Because neither the postfix array nor the stack can ever wind up with more tokens than were
		// contained in the original infix array, only the infix array need be checked for overflow:
		if (infix_count > MAX_TOKENS - 1) // No room for this operator or operand to be added.
			goto abnormal_end;

		map_item &this_map_item = map[map_index];

		switch (this_map_item.type)
		{
		case EXP_DEREF_VAR:
		case EXP_DEREF_FUNC:
		case EXP_DEREF_SINGLE:
			if (infix_count && IS_OPERAND_OR_CPAREN(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
			{
				if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
					goto abnormal_end;
				infix[infix_count++].symbol = SYM_CONCAT;
			}
			switch (this_map_item.type)
			{
			case EXP_DEREF_VAR: // Type() is always VAR_NORMAL as verified earlier. This is relied upon in several places such as built-in functions.
				infix[infix_count].symbol = SYM_VAR; // DllCall() and possibly others rely on this having been done to support changing the value of a parameter (similar to by-ref).
				infix[infix_count].var = this_map_item.var;
				break;
			case EXP_DEREF_FUNC:
				infix[infix_count].symbol = SYM_FUNC;
				infix[infix_count].deref = this_map_item.deref;
				break;
			default: // EXP_DEREF_SINGLE
				// At this stage, an EXP_DEREF_SINGLE item is seen as a numeric literal or a string-literal
				// (without enclosing double quotes, since those are only needed for raw string literals).
				// An EXP_DEREF_SINGLE item cannot extend beyond into the map item to its right, since such
				// a condition can never occur due to load-time preparsing (e.g. the x and y in x+y are two
				// separate items because there's an operator between them). Even a concat expression such as
				// (x y) would still have x and y separate because the space between them counts as a raw
				// map item, which keeps them separate.
				infix[infix_count].symbol = SYM_OPERAND; // Generic string so that it can later be interpreted as a number (if it's numeric).
				infix[infix_count].marker = this_map_item.marker; // This operand has already been terminated above.
			}
			// This map item has been fully processed.  A new loop iteration will be started to move onto
			// the next, if any:
			++infix_count;
			continue;
		}
		
		// Since the above didn't continue, it's either DOUBLE or RAW.

		// An EXP_DEREF_DOUBLE item must be a isolated double-reference or one that extends to the right
		// into other map item(s).  If not, a previous iteration would have merged it in with a previous
		// EXP_RAW item and we could never reach this point.
		// At this stage, an EXP_DEREF_DOUBLE looks like one of the following: abc, 33, abcArray (via
		// extending into an item to its right), or 33Array (overlap).  It can also consist more than
		// two adjacent items as in this example: %ArrayName%[%i%][%j%].  That example would appear as
		// MyArray[33][44] here because the first dereferences have already been done.  MyArray[33][44]
		// (and all the other examples here) are not yet operands because the need a second dereference
		// to resolve them into a number or string.
		if (this_map_item.type == EXP_DEREF_DOUBLE)
		{
			// Find the end of this operand.  StrChrAny() is not used because if *op_end is '\0'
			// (i.e. this_map_item is the last operand), the strchr() below will find that too:
			for (op_end = this_map_item.marker; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
			// Note that above has deteremined op_end correctly because any expression, even those not
			// properly formatted, will have an operator or whitespace between each operand and the next.
			// In the following example, let's say var contains the string -3:
			// %Index%Array var
			// The whitespace-char between the two operands above is a member of EXPR_OPERAND_TERMINATORS,
			// so it (and not the minus inside "var") marks the end of the first operand. If there were no
			// space, the entire thing would be one operand so it wouldn't matter (though in this case, it
			// would form an invalid var-name since dashes can't exist in them, which is caught later).
			cp = this_map_item.marker; // Set for use by the label below.
			goto double_deref;
		}

		// RAW is of lower precedence than the above, so is checked last.  For example, if a single
		// or double deref's contents contain double quotes, those quotes do not delimit a string literal.
		// Instead, the quotes themselves are part of the string.  Similarly, a single or double
		// deref containing as string such as 5+3 is a string, not a subexpression to be evaluated.
		// Since the above didn't "goto" or "continue", this map item is EXP_RAW, which is the only type
		// that can contain operators and raw literal numbers and strings (which are double-quoted when raw).
		for (cp = this_map_item.marker;; ++infix_count) // For each token inside this map item.
		{
			// Because neither the postfix array nor the stack can ever wind up with more tokens than were
			// contained in the original infix array, only the infix array need be checked for overflow:
			if (infix_count > MAX_TOKENS - 1) // No room for this operator or operand to be added.
				goto abnormal_end;

			// Only spaces and tabs are considered whitespace, leaving newlines and other whitespace characters
			// for possible future use:
			cp = omit_leading_whitespace(cp);
			if (cp >= this_map_item.end)
				break; // End of map item (or entire expression if this is the last map item) has been reached.

			terminate_string_here = cp; // See comments below, near other uses of terminate_string_here.

			ExprTokenType &this_infix_item = infix[infix_count]; // Might help reduce code size since it's referenced many places below.

			// Check if it's an operator.
			switch (*cp)
			{
			// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
			case '+':
				sym_prev = infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN; // OPARAN is just a placeholder.
				if (IS_OPERAND_OR_CPAREN(sym_prev)) // CPAREN also covers the tail end of a function call.
					this_infix_item.symbol = SYM_PLUS;
				else // Remove unary pluses from consideration since they do not change the calculation.
					--infix_count; // Counteract the loop's increment.
				break;
			case '-':
				sym_prev = infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN; // OPARAN is just a placeholder.
				// Must allow consecutive unary minuses because otherwise, the following example
				// would not work correctly when y contains a negative value: var := 3 * -y
				if (sym_prev == SYM_NEGATIVE) // Have this negative cancel out the previous negative.
					infix_count -= 2;  // Subtracts 1 for the loop's increment, and 1 to remove the previous item.
				else // Differentiate between unary minus and the "subtract" operator:
				{
					if (IS_OPERAND_OR_CPAREN(sym_prev))
						this_infix_item.symbol = SYM_MINUS;
					else // Unary minus.
					{
						// Set default for cases where the processing below this line doesn't determine
						// it's a negative numeric literal:
						this_infix_item.symbol = SYM_NEGATIVE;
						// v1.0.40.06: The smallest signed 64-bit number (-0x8000000000000000) wasn't properly
						// supported in previous versions because its unary minus was being seen as an operator,
						// and thus the raw number was being passed as a positive to _atoi64() or _strtoi64(),
						// neither of which would recognize it as a valid value.  To correct this, a unary
						// minus followed by a raw numeric literal is now treated as a single literal number
						// rather than unary minus operator followed by a positive number.
						//
						// To be a valid "literal negative number", the character immediately following
						// the unary minus must not be:
						// 1) Whitespace (atoi() and such don't support it, nor is it at all conventional).
						// 2) An open-parenthesis such as the one in -(x).
						// 3) Another unary minus or operator such as --2 (which should evaluate to 2).
						// To cover the above and possibly other unforeseen things, insist that the first
						// character be a digit (even a hex literal must start with 0).
						if (cp[1] >= '0' && cp[1] <= '9')
						{
							// Find the end of this number (this also sets op_end correctly for use by
							// "goto numeric_literal"):
							for (op_end = cp + 2; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
							if (op_end < this_map_item.end) // Detect numeric double derefs such as one created via "12%i% = value".
							{
								// Because the power operator takes precedence over unary minus, don't collapse
								// unary minus into a literal numeric literal if the number is immediately
								// followed by the power operator.  This is correct behavior even for
								// -0x8000000000000000 because -0x8000000000000000**2 would in fact be undefined
								// because +0x8000000000000000 is beyond the signed 64-bit range.
								// Use a temp variable because numeric_literal requires that op_end be set properly:
								char *pow_temp = omit_leading_whitespace(op_end);
								if (!(pow_temp[0] == '*' && pow_temp[1] == '*'))
									goto numeric_literal; // Goto is used for performance and also as a patch to minimize the chance of breaking other things via redesign.
								//else leave this unary minus as an operator.
							}
							//else possible double deref, so leave this unary minus as an operator.
						}
					}
				}
				break;
			case ',':
				this_infix_item.symbol = SYM_COMMA; // It's serves only as a "do not auto-concatenate" indicator for later below.
				break;
			case '/':
				if (cp[1] == '/')
				{
					++cp; // An additional increment to have loop skip over the second '/' too.
					this_infix_item.symbol = SYM_FLOORDIVIDE;
				}
				else
					this_infix_item.symbol = SYM_DIVIDE;
				break;
			case '*':
				if (cp[1] == '*') // Python, Perl, and other languages also use ** for power.
				{
					++cp; // An additional increment to have loop skip over the second '*' too.
					this_infix_item.symbol = SYM_POWER;
				}
				else
				{
					// Differentiate between unary dereference (*) and the "multiply" operator:
					// See '-' above for more details:
					this_infix_item.symbol = IS_OPERAND_OR_CPAREN(infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN)
						? SYM_TIMES : SYM_DEREF;
				}
				break;
			case '!':
				if (cp[1] == '=') // i.e. != is synonymous with <>, which is also already supported by legacy.
				{
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_NOTEQUAL;
				}
				else
					// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
					// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
					// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
					//    always be seen as the binary subtract operator in such cases.
					// 3) Simplifies the code.
					this_infix_item.symbol = SYM_HIGHNOT; // High-precedence counterpart of the word "not".
				break;
			case '(':
				// The below should not hurt any future type-casting feature because the type-cast can be checked
				// for prior to checking the below.  For example, if what immediately follows the open-paren is
				// the string "int)", this symbol is not open-paren at all but instead the unary type-cast-to-int
				// operator.
				if (infix_count && IS_OPERAND_OR_CPAREN(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
						goto abnormal_end;
					this_infix_item.symbol = SYM_CONCAT;
					++infix_count;
				}
				infix[infix_count].symbol = SYM_OPAREN; // Must not refer to "this_infix_item" in case the above did ++infix_count.
				break;
			case ')':
				this_infix_item.symbol = SYM_CPAREN;
				break;
			case '=':
				if (cp[1] == '=')
				{
					// In this case, it's not necessary to check cp >= this_map_item.end prior to ++cp,
					// since symbols such as > and = can't appear in a double-deref, which at
					// this stage must be a legal variable name:
					++cp; // An additional increment to have loop skip over the other '=' too.
					this_infix_item.symbol = SYM_EQUALCASE;
				}
				else
					this_infix_item.symbol = SYM_EQUAL;
				break;
			case '>':
				switch (cp[1])
				{
				case '=':
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_GTOE;
					break;
				case '>':
					++cp; // An additional increment to have loop skip over the second '>' too.
					this_infix_item.symbol = SYM_BITSHIFTRIGHT;
					break;
				default:
					this_infix_item.symbol = SYM_GT;
				}
				break;
			case '<':
				switch (cp[1])
				{
				case '=':
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_LTOE;
					break;
				case '>':
					++cp; // An additional increment to have loop skip over the '>' too.
					this_infix_item.symbol = SYM_NOTEQUAL;
					break;
				case '<':
					++cp; // An additional increment to have loop skip over the second '<' too.
					this_infix_item.symbol = SYM_BITSHIFTLEFT;
					break;
				default:
					this_infix_item.symbol = SYM_LT;
				}
				break;
			case '&':
				if (cp[1] == '&')
				{
					++cp; // An additional increment to have loop skip over the second '&' too.
					this_infix_item.symbol = SYM_AND;
				}
				else
				{
					// Differentiate between unary "take the address of" and the "bitwise and" operator:
					// See '-' above for more details:
					this_infix_item.symbol = IS_OPERAND_OR_CPAREN(infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN)
						? SYM_BITAND : SYM_ADDRESS;
				}
				break;
			case '|':
				if (cp[1] == '|')
				{
					++cp; // An additional increment to have loop skip over the second '|' too.
					this_infix_item.symbol = SYM_OR;
				}
				else
					this_infix_item.symbol = SYM_BITOR;
				break;
			case '^':
				this_infix_item.symbol = SYM_BITXOR;
				break;
			case '~':
				// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
				// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
				// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
				//    always be seen as the binary subtract operator in such cases.
				// 3) Simplifies the code.
				this_infix_item.symbol = SYM_BITNOT;
				break;

			case '"': // Raw string literal.
				// Note that single and double-derefs are impossible inside string-literals
				// because the load-time deref parser would never detect anything inside
				// of quotes -- even non-escaped percent signs -- as derefs.
				// Find the end of this string literal, noting that a pair of double quotes is
				// a literal double quote inside the string:
				++cp; // Omit the starting-quote from consideration, and from the operand's eventual contents.
				for (op_end = cp;; ++op_end)
				{
					if (!*op_end) // No matching end-quote. Probably impossible due to load-time validation.
						goto abnormal_end;
					if (*op_end == '"') // If not followed immediately by another, this is the end of it.
					{
						++op_end;
						if (*op_end != '"') // String terminator or some non-quote character.
							break;  // The previous char is the ending quote.
						//else a pair of quotes, which resolves to a single literal quote.
						// This pair is skipped over and the loop continues until the real end-quote is found.
					}
				}
				// op_end is now the character after the first literal string's ending quote, which might be the terminator.
				*(--op_end) = '\0'; // Remove the ending quote.
				// Convert all pairs of quotes inside into single literal quotes:
				literal_string_length = op_end - cp;
				StrReplace(cp, "\"\"", "\"", SCS_SENSITIVE, UINT_MAX, -1, NULL, &literal_string_length); // PERFORMANCE: Combining this with the loop above isn't trival due to the fact that the terminator isn't added until after that loop.  Another loop could be done here, but it's unlikely to improve perf. much since literal strings tend to be short.
				// Above relies on the fact that StrReplace() does not do cascading replacements,
				// meaning that a series of characters such as """" would be correctly converted into
				// two double quotes rather than just collapsing into one.
				// For the below, see comments at (this_map_item.type == EXP_DEREF_SINGLE) above.
				if (infix_count && IS_OPERAND_OR_CPAREN(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
						goto abnormal_end;
					this_infix_item.symbol = SYM_CONCAT;
					++infix_count;
				}
				// Must not refer to "this_infix_item" in case the above did ++infix_count.
				infix[infix_count].symbol = SYM_STRING; // Marked explicitly as string vs. SYM_OPERAND to prevent it from being seen as a number, e.g. if (var == "12.0") would be false if var has no decimal point.
				infix[infix_count].marker = cp; // This string-operand has already been terminated above.
				cp = op_end + 1;  // Set it up for the next iteration (terminate_string_here is not needed in this case).
				continue;

			default: // Numeric-literal, relational operator such as and/or/not, or unrecognized symbol.
				// Unrecognized symbols should be impossible at this stage because load-time validation
				// would have caught them.  Also, a non-pure-numeric operand should also be impossible
				// because string-literals were handled above, and the load-time validator would not
				// have let any raw non-numeric operands get this far (such operands would have been
				// converted to single or double derefs at load-time, in which case they wouldn't be
				// raw and would never reach this point in the code).
				// To conform to the way the load-time pre-parser recognizes and/or/not, and to support
				// things like (x=3)and(5=4) or even "x and!y", the and/or/not operators are processed
				// here with the numeric literals since we want to find op_end the same way.

				if (*cp == '.' && IS_SPACE_OR_TAB(cp[1])) // This one must be done here rather than as a "case".  See comment below.
				{
					this_infix_item.symbol = SYM_CONCAT;
					break;
				}
				//else any '.' not followed by a space or tab is likely a number without a leading zero,
				// so continue on below to process it.

				// Find the end of this operand or keyword, even if that end is beyond this_map_item.end.
				// StrChrAny() is not used because if *op_end is '\0', the strchr() below will find it too:
				for (op_end = cp + 1; !strchr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
numeric_literal:
				// Now op_end marks the end of this operand or keyword.  That end might be the zero terminator
				// or the next operator in the expression, or just a whitespace.
				if (op_end >= this_map_item.end // This must be true to qualify as a double deref.
					&& (*this_map_item.end // If this is true, it's enough to know that it's a double deref.
					// But if not, and all three of the following are true, it's a double deref anyway
					// to support the correct result in something like: Var := "x" . Array%BlankVar%
					|| (map_index != map_count - 1 && MAP_ITEM_IN_BUFFER(map[map_index + 1].type)
						&& map[map_index + 1].marker == op_end)))
					goto double_deref; // This also serves to break out of this for(), equivalent to a break.
				// Otherwise, this operand is a normal raw numeric-literal or a word-operator (and/or/not).
				// The section below is very similar to the one used at load-time to recognize and/or/not,
				// so it should be maintained with that section.  UPDATE for v1.0.45: The load-time parser
				// now resolves "OR" to || and "AND" to && to improve runtime performance and reduce code size here.
				// However, "NOT" but still be parsed here at runtime because it's not quite the same as the "!"
				// operator (different precedence), and it seemed too much trouble to invent some special
				// operator symbol for load-time to insert as a placeholder/substitute (especially since that
				// symbol would appear in ListLines).
				if (op_end-cp == 3
					&& (cp[0] == 'n' || cp[0] == 'N')
					&& (cp[1] == 'o' || cp[1] == 'O')
					&& (cp[2] == 't' || cp[2] == 'T')) // "NOT" was found.
				{
					this_infix_item.symbol = SYM_LOWNOT;
					*cp = '\0';  // Terminate any previous raw numeric-literal such as "1 not (x < 3)" (even though "not" would be invalid if used this way)
					cp = op_end; // Have the loop process whatever lies at op_end and beyond.
					continue;
				}
				// Since above didn't "continue", this item is a raw numeric literal, either SYM_FLOAT or
				// SYM_INTEGER (to be differentiated later).
				// For the below, see comments at (this_map_item.type == EXP_DEREF_SINGLE) above.
				if (infix_count && IS_OPERAND_OR_CPAREN(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
						goto abnormal_end;
					this_infix_item.symbol = SYM_CONCAT;
					++infix_count;
				}
				// Must not refer to "this_infix_item" in case above did ++infix_count:
				infix[infix_count].symbol = SYM_OPERAND;
				infix[infix_count].marker = cp; // This numeric operand will be terminated later via terminate_string_here.
				cp = op_end; // Have the loop process whatever lies at op_end and beyond.
				// The below is necessary to support an expression such as (1 "" 0), which
				// would otherwise result in 1"0 instead of 10 because the 1 was lazily
				// terminated by the next iteration rather than our iteration at it's
				// precise viewed-as-string ending point.  It might also be needed for
				// the same reason for concatenating things like (1 var).
				if (IS_SPACE_OR_TAB(*cp))
					*cp++ = '\0';
				continue; // i.e. don't do the terminate_string_here and ++cp steps below.
			} // switch() for type of symbol/operand.

			// If the above didn't "continue", it just processed a non-operand symbol.  So terminate
			// the string at the first character of that symbol (e.g. the first character of <=).
			// This sets up raw operands to be always-terminated, such as the ones in 5+10+20.  Note
			// that this is not done for operator-words (and/or/not) since it's not valid to write
			// something like 1and3 (such a thing would be considered a variable and converted into
			// a single-deref by the load-time pre-parser).  It's done this way because we don't
			// want to convert these raw operands into numbers yet because their original strings
			// might be needed in the case where this operand will be involved in an operation with
			// another string operand, in which case both are treated as strings:
			if (terminate_string_here)
				*terminate_string_here = '\0';
			++cp; // i.e. increment only if a "continue" wasn't encountered somewhere above.
		} // for each token
		continue;  // To avoid falling into the label below.

double_deref:
		// The only purpose of the following loop is to increase map_index if one or more of the map items
		// to the right of this_map_item are to be merged with this_map_item to construct a double deref
		// such as Array%i%.
		for (++map_index;; ++map_index)
		{
			if (map_index == map_count || !MAP_ITEM_IN_BUFFER(map[map_index].type)
				|| (op_end <= map[map_index].marker  // Since above line didn't short-circuit, it's safe to reference map[map_index].marker.
					&& map[map_index].end > map[map_index].marker))
				// The final line above serves to merge empty items (which must be doubles since RAWs are never
				// empty) in with this one.  Although everything might work correctly without this, it's more
				// proper to get rid of these empty items now since they should "belong" to this item.
			{
				// The map item to the right of the one containg the end of this operand has been found.
				--map_index;
				// If the loop had only one iteration, the above restores the original value of map_index.
				// In other words, this map item doesn't stretch into others to its right, so it's just a
				// naked double-deref such as %DynVar%.
				break;
			}
		}
		// If the map_item[map_index] item isn't fully consumed by this operand, alter the map item to contain
		// only the part left to be processed and then have the loop process this same map item again.
		// For example, in Array[%i%]/3, the final map item is ]/3, of which only the ] is consumed by the
		// Array[%i%] operand.
		if (op_end < map[map_index].end)
		{
			if (map[map_index].type == EXP_RAW)
			{
				map[map_index].marker = op_end;
				--map_index;  // Compensate for the loop's ++map_index.
			}
			else // DOUBLE or something else that shouldn't be allowed to be partially processed such as the above.
			{
				// The above EXP_RAW method is not done of the map item is a double deref, since it's not
				// currently valid to do something like Var:=%VarContainingSpaces% + 1.  Example:
				//var = test
				//x = var 11
				//y := %x% + 1  ; Add 1 to force it to stay an expression rather than getting simplified at loadtime.
				// In such cases, force it to handle this entire double as a unit, since other usages are invalid.
				op_end = map[map_index].end; // Force the entire map item to be processed/consumed here.
			}
		}
		//else do nothing since map_index is now set to the final map item of this operand, and that
		// map item is fully consumed by this operand and needs no further processing.

		// UPDATE: The following is now supported in v1.0.31, so this old comment is kept only for background:
		// Check if this double is being concatenated onto a previous operand.  If so, it is not
		// currently supported so this double-deref will be treated as an empty string, as documented.
		// Example 1: Var := "abc" %naked_double_ref%
		// Example 2: Var := "abc" Array%Index%
		// UPDATE: Here is the means by which the above is now supported:
		if (infix_count && IS_OPERAND_OR_CPAREN(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
		{
			if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
				goto abnormal_end;
			infix[infix_count++].symbol = SYM_CONCAT;
		}

		if (   !(op_length = (UINT)(op_end - cp))   )
		{
			// Var is not found, not a normal var, or it *is* an environment variable.
			infix[infix_count].symbol = SYM_OPERAND;
			infix[infix_count].marker = "";
		}
		else // This operand becomes the variable's contents.
		{
			// Callers of this label have set cp to the start of the variable name and op_end to the
			// position of the character after the last one in the name.
			// In v1.0.31, FindOrAddVar() vs. FindVar() is called below to support the passing of non-existent
			// array elements ByRef, e.g. Var:=MyFunc(Array%i%) where the MyFunc function's parameter is
			// defined as ByRef, would effectively create the new element Array%i% if it doesn't already exist.
			// Since at this stage we don't know whether this particular double deref is to be sent as a param
			// to a function, or whether it will be byref, this is done unconditionally for all double derefs
			// since it seems relatively harmless to create a blank variable in something like var := Array%i%
			// (though it will produce a runtime error if the double resolves to an illegal variable name such
			// as one containing spaces).
			// The use of ALWAYS_PREFER_LOCAL below improves flexibility of assume-global functions
			// by allowing this command to resolve to a local first if such a local exists:
			if (   !(found_var = g_script.FindOrAddVar(cp, op_length, ALWAYS_PREFER_LOCAL))   ) // i.e. don't call FindOrAddVar with zero for length, since that's a special mode.
			{
				// Above already displayed the error.  As of v1.0.31, this type of error is displayed and
				// causes the current thread to terminate, which seems more useful than the old behavior
				// that tolerated anything in expressions.
				goto abort;
			}
			// Otherwise, var was found or created.
			if (found_var->Type() != VAR_NORMAL)
			{
				// Non-normal variables such as Clipboard and A_ScriptFullPath are not allowed to be
				// generated from a double-deref such as A_Script%VarContainingFullPath% because:
				// 1) Anything that needed their contents would have to find memory in which to store
				//    the result of Var::Get(), which would complicate the code since that would have
				//    to be added 
				// 2) It doesn't appear to have much use, not even for passing them as a ByRef parameter
				// to a function (since they're read-only [except Clipboard, but temporary memory would be
				// needed somewhere if the clipboard contains files that need to be expanded to text] and
				// essentially global by their very nature), and the value of catching unintended usages
				// seems more important than any flexibilty that might add.
				infix[infix_count].symbol = SYM_OPERAND;
				infix[infix_count].marker = "";
			}
			else
			{
				// Even if it's an environment variable, it gets added as SYM_VAR.  However, unlike other
				// aspects of the program, double-derefs that resolve to environment variables will be seen
				// as always-blank due to the use of Var::Contents() vs. Var::Get() in various places below.
				// This seems okay due to the extreme rarity of anyone intentionally wanting a double
				// reference such as Array%i% to resolve to the name of an environment variable.
				infix[infix_count].symbol = SYM_VAR; // The fact that a SYM_VAR operand is always VAR_NORMAL is relied upon in several places such as built-in functions.
				infix[infix_count].var = found_var;
			}
		} // Double-deref section.
		++infix_count;
	} // for each map item

	////////////////////////////
	// CONVERT INFIX TO POSTFIX.
	////////////////////////////
	#define STACK_PUSH(token) stack[stack_count++] = &token
	#define STACK_POP stack[--stack_count]  // To be used as the r-value for an assignment.
	// SYM_BEGIN is the first item to go on the stack.  It's a flag to indicate that conversion to postfix has begun:
	ExprTokenType token_begin;
	token_begin.symbol = SYM_BEGIN;
	STACK_PUSH(token_begin);

	int i;
	SymbolType stack_symbol, infix_symbol;

	for (i = 0; stack_count > 0;) // While SYM_BEGIN is still on the stack, continue iterating.
	{
		stack_symbol = stack[stack_count - 1]->symbol; // Frequently used, so resolve only once to help performance.

		// "i" will be out of bounds if infix expression is complete but stack is not empty.
		// So the very first check must be for that.
		if (i == infix_count) // End of infix expression, but loop's check says stack still has items on it.
		{
			if (stack_symbol == SYM_BEGIN) // Stack is basically empty, so stop the loop.
				// Remove SYM_BEGIN from the stack, leaving the stack empty for use in the next stage.
				// This also signals our loop to stop.
				--stack_count;
			else if (stack_symbol == SYM_OPAREN) // Open paren is never closed (currently impossible due to load-time balancing, but kept for completeness).
				goto abnormal_end;
			else // Pop item of the stack, and continue iterating, which will hit this line until stack is empty.
			{
				postfix[postfix_count] = STACK_POP;
				postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
			}
			continue;
		}

		// Only after the above is it safe to use "i" as an index.
		infix_symbol = infix[i].symbol; // Frequently used, so resolve only once to help performance.

		// Put operands into the postfix array immediately, then move on to the next infix item:
		if (IS_OPERAND(infix_symbol)) // At this stage, operands consist of only SYM_OPERAND and SYM_STRING.
		{
			postfix[postfix_count] = infix + i++;
			postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
			continue;
		}

		// Since above didn't "continue", the current infix symbol is not an operand.
		switch(infix_symbol)
		{
		// CPAREN is listed first for performance.  It occurs frequently while emptying the stack to search
		// for the matching open-paren:
		case SYM_CPAREN:
			if (stack_symbol == SYM_OPAREN) // The first open-paren on the stack must be the one that goes with this close-paren.
			{
				--stack_count; // Remove this open-paren from the stack, since it is now complete.
				++i;           // Since this pair of parentheses is done, move on to the next token in the infix expression.
				// There should be no danger of stack underflow in the following because SYM_BEGIN always
				// exists at the bottom of the stack:
				if (stack[stack_count - 1]->symbol == SYM_FUNC) // Within the postfix list, a function-call should always immediately follow its params.
				{
					postfix[postfix_count] = STACK_POP;
					postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
				}
			}
			else if (stack_symbol == SYM_BEGIN) // Paren is closed without having been opened (currently impossible due to load-time balancing, but kept for completeness).
				goto abnormal_end; 
			else
			{
				postfix[postfix_count] = STACK_POP;
				postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
				// By not incrementing i, the loop will continue to encounter SYM_CPAREN and thus
				// continue to pop things off the stack until the corresponding OPAREN is reached.
			}
			break;

		// Open-parentheses always go on the stack to await their matching close-parentheses:
		case SYM_OPAREN:
			STACK_PUSH(infix[i++]);
			break;

		case SYM_COMMA:
			// Fix for v1.0.31.01: Commas must force everything off the stack until this comma's own function
			// call is encountered on the stack.  Otherwise, an expression such as fn(a+b, c) would be incorrectly
			// converted to postfix "a b c + fn()" (i.e. the plus would operate upon b & c rather than a & b).
			// First function-call on the stack must own this comma if expression is syntactically correct.
			// Each function-call is accompanied by its open-parenthesis on the stack:
			if (stack_symbol != SYM_OPAREN || stack[stack_count - 2]->symbol != SYM_FUNC) // Relies on short-circuit boolean order.
			{
				postfix[postfix_count] = STACK_POP;
				postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
				// And by not incrementing i, this comma/case will continue to be encountered until everything comes off the stack the needs to be.
			}
			else
				++i; // Omit commas from further consideration, since they only served as a "do not concatenate" indicator earlier.
			break;

		default: // Symbol is an operator, so act according to its precedence.
			// If the symbol waiting on the stack has a lower precedence than the current symbol, push the
			// current symbol onto the stack so that it will be processed sooner than the waiting one.
			// Otherwise, pop waiting items off the stack (by means of i not being incremented) until their
			// precedence falls below the current item's precedence, or the stack is emptied.
			// Note: BEGIN and OPAREN are the lowest precedence items ever to appear on the stack (CPAREN
			// never goes on the stack, so can't be encountered there).
			if (   sPrecedence[stack_symbol] < sPrecedence[infix_symbol]
				|| stack_symbol == SYM_POWER && (infix_symbol >= SYM_NEGATIVE && infix_symbol <= SYM_ADDRESS
					|| infix_symbol == SYM_LOWNOT)   )
			{
				// The line above is a workaround to allow 2**-2 (and others in v1.0.45) to be evaluated as
				// 2**(-2) rather than being seen as an error.  However, for simplicity of code, consecutive
				// unary operators are not supported (they currently produce a failure [blank value]
				// because they wind up in the postfix array in the wrong order).
				// !!x  ; Not supported, but Laszlo pointed out this would useful to convert a blank value into
				//      ; a zero for use with unitialized variables.  But the workaround is to use !(!x).
				//      ; A lesser reason not to support it is to reserve !! for future use as a new operator,
				//      ; though that may be too unlikely due to its counterintuitiveness.
				//      ; Similarly, "not not x" is not supported for the same reasons as the others here.
				// !-3  ; Not supported (seems of little use anyway; can be written as !(-3) to make it work).
				// -!3  ; Not supported (seems useless anyway, can be written as -(!3) to make it work).
				// !x   ; Supported even if X contains a negative number, since x is recognized as an isolated operand and not something containing unary minus.
				// !&Var ; Not supported (seems useless anyway; can be written with parentheses to make it work).
				// -&Var ; Same
				// ~&Var ; Same
				// !*Var, -*Var and ~*Var: These are supported by means of having * be a higher precedence than the other unary operators.

				// To facilitate short-circuit boolean evaluation, right before an AND/OR is pushed onto the
				// stack, connect the end of it's left branch to it.  Note that the following postfix token
				// can itself be of type AND/OR, a simple example of which is "if (true and true and true)",
				// in which the first and's parent in an imaginary tree is the second "and".
				// But how is it certain that this is the final operator or operand of and AND/OR's left branch?
				// Here is the explanation:
				// Everything higher priority than the AND/OR came off the stack right before it, resulting in
				// what must be a balanced/complete sub-postfix-expression in and of itself (unless the expression
				// has a syntax error, which is caught in various places).  Because it's complete, during the
				// postfix evaluation phase, that sub-expression will result in a new operand for the stack,
				// which must then be the left side of the AND/OR because the right side immediately follows it
				// within the postfix array, which in turn is immediately followed its operator (namely AND/OR).
				if ((infix_symbol == SYM_AND || infix_symbol == SYM_OR) && postfix_count)
					postfix[postfix_count - 1]->circuit_token = infix + i;
				STACK_PUSH(infix[i++]);
			}
			else // Stack item has equal or greater precedence (if equal, left-to-right evaluation order is in effect).
			{
				postfix[postfix_count] = STACK_POP;
				postfix[postfix_count++]->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
			}
		} // switch(infix_symbol)
	} // End of loop that builds postfix array.

	///////////////////////////////////////////////////
	// EVALUATE POSTFIX EXPRESSION (constructed above).
	///////////////////////////////////////////////////
	SymbolType right_is_number, left_is_number;
	double right_double, left_double;
	__int64 right_int64, left_int64;
	char *right_string, *left_string;
	char *right_contents, *left_contents;
	size_t right_length, left_length;
	char left_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];  // BIF_OnMessage relies on this one being large enough to hold MAX_VAR_NAME_LENGTH.
	char right_buf[MAX_FORMATTED_NUMBER_LENGTH + 1]; // Only needed for holding numbers
	int j, s, actual_param_count;
	Func *prev_func;
	char *result; // "result" is used for return values and also the final result.
	size_t result_size, alloca_usage = 0; // v1.0.45: Track amount of alloca mem to avoid stress on stack from extreme expressions (mostly theoretical).
	bool done, udf_is_final_action, make_result_persistent, early_return, left_branch_is_true, left_was_negative;
	ExprTokenType *circuit_token;
	VarBkp *var_backup = NULL;  // If needed, it will hold an array of VarBkp objects. v1.0.40.07: Initialized to NULL to facilitate an approach that's more maintainable.
	int var_backup_count; // The number of items in the above array (when it's non-NULL).

	// For each item in the postfix array: if it's an operand, push it onto stack; if it's an operator or
	// function call, evaluate it and push its result onto the stack.
	for (i = 0; i < postfix_count; ++i)
	{
		ExprTokenType &this_token = *postfix[i];  // For performance and convenience.

		// At this stage, operands in the postfix array should be either SYM_OPERAND or SYM_STRING.
		// But all are checked since that operation is just as fast:
		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
			goto push_this_token;

		if (this_token.symbol == SYM_FUNC) // A call to a function (either built-in or defined by the script).
		{
			Func &func = *this_token.deref->func; // For performance.
			actual_param_count = this_token.deref->param_count; // For performance.
			if (actual_param_count > stack_count) // Prevent stack underflow (probably impossible if actual_param_count is accurate).
				goto abnormal_end;
			if (func.mIsBuiltIn)
			{
				// Adjust the stack early to simplify.  Above already confirmed that this won't underflow.
				// Pop the actual number of params involved in this function-call off the stack.  Load-time
				// validation has ensured that this number is always less than or equal to the number of
				// parameters formally defined by the function.  Therefore, there should never be any leftover
				// function-params on the stack after this is done:
				stack_count -= actual_param_count; // The function called below will see this portion of the stack as an array of its parameters.
				this_token.symbol = SYM_INTEGER; // Set default return type so that functions don't have to do it if they return INTs.
				this_token.marker = func.mName;  // Inform function of which built-in function called it (allows code sharing/reduction). Can't use circuit_token because it's value is still needed later below.
				this_token.buf = left_buf;       // mBIF() can use this to store a string result, and for other purposes.

				// BACKUP THE CIRCUIT TOKEN:
				circuit_token = this_token.circuit_token; // Backup the current circuit_token, which can be non-NULL (verified through code review).
				this_token.circuit_token = NULL; // Init to detect whether the called function allocates it (i.e. we're overloading it with a new purpose).

				// CALL THE FUNCTION:
				func.mBIF(this_token, stack + stack_count, actual_param_count);

				// RESTORE IT (after handling what came back inside it):
				#define EXPR_IS_DONE (!stack_count && i == postfix_count - 1) // True if we've used up the last of the operators & operands.
				done = output_var && EXPR_IS_DONE; // i.e. this is ACT_ASSIGNEXPR and we've now produced the final result.
				if (this_token.circuit_token) // The called function allocated some memory here and turned it over to us.
				{
					// In most cases, the string stored in circuit_token is the same address as this_token.marker
					// (i.e. what is named "result" further below), because that's what the built-in functions
					// are normally using the memory for.
					if (done && (char *)this_token.circuit_token == this_token.marker) // circuit_token is checked in case caller alloc'd mem but didn't use it as its actual result.
					{
						// v1.0.45: Take a shortcut for performance.  Doing it this way saves at least two
						// memcpy's (one into deref buffer and then another back into the output_var by
						// ACT_ASSIGNEXPR itself).  In some cases is also saves from having to expand the deref buffer
						// as well as the output_var (since it's current memory might be too small to hold the
						// new memory block). Thus we give it a new block directly to avoid all of that.
						// This should be a big boost to performance when long strings are involved.
						output_var->AcceptNewMem((char *)this_token.circuit_token, (VarSizeType)(size_t)this_token.buf); // The called function is responsible for having stored the length as an overload of this_token.buf.
						goto normal_end_skip_output_var; // No need to restore circuit_token because the expression is finished.
					}
					// Otherwise, not done yet; so handle this memory the normal way: Mark it to be freed at the
					// time we return.
					if (mem_count == MAX_EXPR_MEM_ITEMS) // No more slots left (should be nearly impossible).
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
					mem[mem_count++] = (char *)this_token.circuit_token;
				}
				this_token.circuit_token = circuit_token; // Restore it to its original value.

				// HANDLE THE RESULT (unless it was already handled above due to an optimization):
				if (IS_NUMERIC(this_token.symbol)) // Any numeric result can be considered final because it's already stored in permanent memory (namely the token itself).
					goto push_this_token; // For code simplicity, the optimization for numeric results is done at a later stage.
				//else it's a string, which might need to be moved to persistent memory further below.
				if (done) // In this case, "done" means the expression is finished *and* we have an output_var available.
				{
					// v1.0.45: This mode improves performance by avoiding the need to copy the result into
					// more persistent memory, then avoiding the need to copy it into the defer buffer (which
					// also avoids the possibility of needing to expand that buffer).
					output_var->Assign(this_token.marker); // Marker can be used because symbol will never be SYM_VAR
					goto normal_end_skip_output_var;       // in this case. ALSO: Assign() contains an optimization that avoids actually doing the mem-copying if output_var is being assigned to itself (which can happen in cases like RegExMatch()).
				}
				// Otherwise, there's no output_var or the expressions isn't finished yet, so do normal processing.
				result = this_token.marker; // Marker can be used because symbol will never be SYM_VAR in this case.
				early_return = false; // For maintainability.  A built-in function can't cause an early return, unlike the Exit command in a user-defined function.
			}
			else // It's not a built-in function, or it's a built-in that was overridden with a custom function.
			{
				udf_is_final_action = false; // Set default.
				// If there are other instances of this function already running, either via recursion or
				// an interrupted quasi-thread, back up the local variables of the instance that lies immediately
				// beneath ours (in turn, that instance is responsible for backing up any instance that lies
				// beneath it, and so on, since when recursion collapses or threads resume, they always do so
				// in the reverse order in which they were created.
				//
				// I think the backup-and-restore approach to local variables might enhance performance over
				// other approaches, perhaps a lot.  This is because most of the time there will be no other
				// instances of a given function on the call stack, thus no backup/restore is needed, and thus
				// the function's existing local variables can be reused as though they're globals (i.e.
				// memory allocation/deallocation overhead is often completely avoided for non-recursive calls
				// to a function after the first).
				if (func.mInstances > 0) // i.e. treat negatives as zero to help catch any bugs in the way mInstances is maintained.
				{
					// Backup/restore of function's variables is needed.
					// Only when a backup is needed is it possible for this function to be calling itself recursively,
					// either directly or indirectly by means of an intermediate function.  As a consequence, it's
					// possible for this function to be passing one or more of its own params or locals to itself.
					// The following section compensates for that to handle parameters passed by-value, but it
					// doesn't correctly handle passing its own locals/params to itself ByRef, which is in the
					// help file as a known limitation.  Also, the below doesn't indicate a failure when stack
					// underflow would occur because the loop after this one needs to do that (since this
					// one will never execute if a backup isn't needed).  Note that this loop that reviews all
					// actual parameters is necessary as a separate loop from the one further below because this
					// first one's conversion must occur prior to calling BackupFunctionVars().  In addition, there
					// might be other interdepencies between formals and actuals if a function is calling itself
					// recursively.
					for (j = func.mParamCount - 1, s = stack_count; j > -1; --j) // For each formal parameter (reverse order to mirror the nature of the stack).
					{
						if (j < actual_param_count) // This formal has an actual on the stack.
						{
							// Move on to the next item in the stack (without popping):  A check higher above
							// has already ensured that this won't cause stack underflow:
							--s;
							if (stack[s]->symbol == SYM_VAR && !func.mParam[j].var->IsByRef())
							{
								// Since this formal parameter is passed by value, if it's SYM_VAR, convert it to
								// SYM_OPERAND to allow the variables to be backed up and reset further below without
								// corrupting any SYM_VARs that happen to be locals or params of this very same
								// function.
								// DllCall() relies on the fact that this transformation is only done for UDFs
								// and not built-in functions such as DllCall().  This is because DllCall() sometimes
								// needs the variable of a parameter for use as an output parameter.
								stack[s]->marker = stack[s]->var->Contents();
								stack[s]->symbol = SYM_OPERAND;
							}
						}
					}
					// BackupFunctionVars() will also clear each local variable and formal parameter so that
					// if that parameter or local var or is assigned a value by any other means during our call
					// to it, new memory will be allocated to hold that value rather than overwriting the
					// underlying recursed/interrupted instance's memory, which it will need intact when it's resumed.
					if (!BackupFunctionVars(func, var_backup, var_backup_count)) // Out of memory.
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
				}
				//else backup is not needed because there are no other instances of this function on the call-stack.
				// So by definition, this function is not calling itself directly or indirectly, therefore there's no
				// need to do the conversion of SYM_VAR because those SYM_VARs can't be ones that were blanked out
				// due to a function exiting.  In other words, it seems impossible for a there to be no other
				// instances of this function on the call-stack and yet SYM_VAR to be one of this function's own
				// locals or formal params because it would have no legitimate origin.

				// Pop the actual number of params involved in this function-call off the stack.  Load-time
				// validation has ensured that this number is always less than or equal to the number of
				// parameters formally defined by the function.  Therefore, there should never be any leftover
				// params on the stack after this is done:
				for (j = func.mParamCount - 1; j > -1; --j) // For each formal parameter (reverse order to mirror the nature of the stack).
				{
					FuncParam &this_formal_param = func.mParam[j]; // For performance and convenience.
					if (j >= actual_param_count) // No actual to go with it (should be possible only if the parameter is optional or has a default value).
					{
						switch(this_formal_param.default_type)
						{
						case PARAM_DEFAULT_STR:
							this_formal_param.var->Assign(this_formal_param.default_str);
							break;
						case PARAM_DEFAULT_INT:
							this_formal_param.var->Assign(this_formal_param.default_int64);
							break;
						case PARAM_DEFAULT_FLOAT:
							this_formal_param.var->Assign(this_formal_param.default_double);
							break;
						default: // PARAM_DEFAULT_NONE or some other value.  This is probably a bug; assign blank for now.
							this_formal_param.var->Assign(); // By not specifying "" as the first param, the var's memory is not freed, which seems best to help performance when the function is called repeatedly in a loop.
							break;
						}
						continue;
					}
					// Otherwise, assign actual parameter's value to the formal parameter (which is itself a
					// local variable in the function).  A check higher above has already ensured that this
					// won't cause stack underflow:
					ExprTokenType &token = *STACK_POP;
					// Below uses IS_OPERAND rather than checking for only SYM_OPERAND because the stack can contain
					// both generic and specific operands.  Specific operands were evaluated by a previous iteration
					// of this section.  Generic ones were pushed as-is onto the stack by a previous iteration.
					if (!IS_OPERAND(token.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
						goto abort_udf;  // Done in lieu of "goto end/fail" so that Free()+RestoreFunctionVars() can be called.
					if (this_formal_param.var->IsByRef())
					{
						// Note that the previous loop might not have checked things like the following because that
						// loop never ran unless a backup was needed:
						if (token.symbol != SYM_VAR)
						{
							// In most cases this condition would have been caught by load-time validation.
							// However, in the case of badly constructed double derefs, that won't be true
							// (though currently, only a double deref that resolves to a built-in variable
							// would be able to get this far to trigger this error, because something like
							// func(Array%VarContainingSpaces%) would have been caught at an earlier stage above.
							LineError(ERR_BYREF ERR_ABORT, FAIL, this_formal_param.var->mName);
							goto abort_udf;  // Done in lieu of "goto end/fail" so that Free()+RestoreFunctionVars() can be called.
						}
						this_formal_param.var->UpdateAlias(token.var); // Make the formal parameter point directly to the actual parameter's contents.
					}
					else // This parameter is passed "by value".
					{
						switch(token.symbol)
						{
						case SYM_VAR:
							// This case can still happen because the previous loop's conversion of all
							// by-value SYM_VAR operands into SYM_OPERAND would not have happened if no
							// backup was needed for this function.
							// v1.0.45: Pass length (since it's known) so that Assign() doesn't have to call strlen().
							// But ignore binary clipboard because it's documented that except for certain features,
							// binary clipboard variables are seen only up to the first binary zero (mostly to
							// simplify the code).
							this_formal_param.var->Assign(token.var->Contents(), token.var->LengthIgnoreBinaryClip());
							break;
						case SYM_INTEGER:
							this_formal_param.var->Assign(token.value_int64);
							break;
						case SYM_FLOAT: // Listed last for performance.
							this_formal_param.var->Assign(token.value_double);
							break;
						default: // SYM_STRING or SYM_OPERAND
							this_formal_param.var->Assign(token.marker);
						}
					}
				}

				result = ""; // Init to default in case function doesn't return a value or it EXITs or fails.

				// Launch the function similar to Gosub (i.e. not as a new quasi-thread):
				// The performance gain of conditionally passing NULL in place of result (when this is the
				// outermost function call of a line consisting only of function calls, namely ACT_FUNCTIONCALL)
				// would not be significant because the Return command's expression (arg1) must still be evaluated
				// in case it calls any functions that have side-effects, e.g. "return LogThisError()".
				prev_func = g.CurrentFunc; // This will be non-NULL when a function is called from inside another function.
				g.CurrentFunc = &func;
				// Although a GOTO that jumps to a position outside of the function's body could be supported,
				// it seems best not to for these reasons:
				// 1) The extreme rarity of a legitimate desire to intentionally do so.
				// 2) The fact that any return encountered after the Goto cannot provide a return value for
				//    the function because load-time validation checks for this (it's preferable not to
				//    give up this check, since it is an informative error message and might also help catch
				//    bugs in the script).  Gosub does not suffer from this because the return that brings it
				//    back into the function body belongs to the Gosub and not the function itself.
				// 3) More difficult to maintain because we have handle jump_to_line the same way ExecUntil() does,
				//    checking aResult the same way it does, then checking jump_to_line the same way it does, etc.
				// Fix for v1.0.31.05: g.mLoopFile and the other g_script members that follow it are
				// now passed to ExecUntil() for two reasons (update for v1.0.44.14: now they're implicitly "passed"
				// because they're done via parameter anymore):
				// 1) To fix the fact that any function call in one parameter of a command would reset
				// A_Index and related variables so that if those variables are referenced in another
				// parameter of the same command, they would be wrong.
				// 2) So that the caller's value of A_Index and such will always be valid even inside
				// of called functions (unless overridden/eclipsed by a loop in the body of the function),
				// which seems to add flexibility without giving up anything.  This fix is necessary at least
				// for a command that references A_Index in two of its args such as the following:
				// ToolTip, O, ((cos(A_Index) * 500) + 500), A_Index
				++func.mInstances;
				aResult = func.mJumpToLine->ExecUntil(UNTIL_BLOCK_END, &result, NULL);
				--func.mInstances;
				// Restore the original value in case this function is called from inside another function.
				// Due to the synchronous nature of recursion and recursion-collapse, this should keep
				// g.CurrentFunc accurate, even amidst the asynchronous saving and restoring of "g" itself:
				g.CurrentFunc = prev_func;

				if (   !(early_return = (aResult == EARLY_EXIT || aResult == FAIL))   ) // No need to do any of this for early_return because for backward compatibility, ACT_ASSIGNEXPR is aborted by early return (i.e. output_var's original contents are not altered).
				{
					if (output_var && EXPR_IS_DONE) // i.e. this is ACT_ASSIGNEXPR and we've now produced the final result.
					{
						// v1.0.45: Take a shortcut for performance.  Doing it this way saves up to two memcpy's
						// (make_result_persistent then copy into deref buffer).  In some cases, it also saves
						// from having to make_result_persistent and prevents the need to expand the deref buffer.
						// HOWEVER, the following optimization isn't done because not only does it complicate the
						// code a lot (such as verifying that the variable isn't static, isn't ALLOC_SIMPLE, isn't
						// a ByRef to a global or some other function's local, etc.), there's also currently
						// no way to find out which function owns a particular local variable (a name lookup
						// via binary search is a possibility, but its performance probably isn't worth it)
						// Abandoned idea: When a user-defined function returns one of its local variables,
						// the contents of that local variable can be salvaged if it's about to be destroyed
						// anyway in conjunction with normal function-call cleanup. In other words, we can take
						// that local variable's memory and directly hang it onto the output_var.
						output_var->Assign(result);
						udf_is_final_action = true; // This tells the label below that after the cleanup, we're done.
						goto skip_abort_udf; // Do end-of-function-call cleanup (see comment above).  No need to do make_result_persistent section.
					}
					// Otherwise (since above didn't goto): !output_var || !EXPR_IS_DONE, so do normal handling of
					// "result" below.
				}
				//else early_return==true, so no need to store anything in result because for backward compatibility, this expression will have no result storable by the outside world (e.g. ACT_ASSIGNEXPR).
			} // Call to a user defined function.

			// Above just called a function (either user-defined or built-in).
			done = EXPR_IS_DONE; // Macro called again because its components may have changed. Resolve macro only once for use in more than one place below.

			// The result just returned needs to be copied to a more persistent location.  This is done right
			// away if the result is the contents of a local variable (since all locals are about to be freed
			// and overwritten), which is assumed to be the case if it's not in the new deref buf because it's
			// difficult to distinguish between when the function returned one of its own local variables
			// rather than a global or a string/numeric literal).  The only exceptions are:
			if (early_return // We're about to return early, so the caller will be ignoring this result entirely.
				|| done && mActionType == ACT_FUNCTIONCALL // Outermost function call's result will be ignored, so no need to store it.
				|| mem_count && result == mem[mem_count - 1]) // v1.0.45: A call to a built-in function can sometimes store its result here, which is already persistent.
				make_result_persistent = false;
			else if (result < sDerefBuf || result >= sDerefBuf + sDerefBufSize) // Not in their deref buffer (yields correct result even if sDerefBuf is NULL).
				make_result_persistent = true; // Since above didn't set it to false, this result must be assumed to be one of their local variables, so must be immediately copied since it's about to be cleared.
			// So now since the above didn't set the value, the result must be in their deref buffer, perhaps
			// due to something like "return x+3" on their part.
			else if (done) // We don't have to make it persistent here because the final stage will copy it from their deref buf into ours (since theirs is only deleted later, by our caller).
				make_result_persistent = false;
			else // There are more operators/operands to be evaluated, but if there are no more function calls, we don't have to make it persistent since their deref buf won't be overwritten by anything during the time we need it.
			{
				if (func.mIsBuiltIn)
					make_result_persistent = true; // Future operators/operands might use the buffer where the result is stored, so must copy it somewhere else (unless "done", which was already checked above).
				else
				{
					make_result_persistent = false; // Set default to be possibly overridden below.
					// Since there's more in the stack or postfix array to be evaluated, and since the return value
					// is in the new deref buffer, must copy result to somewhere non-volatile whenever there's
					// another function call pending by us.  But if result is the empty string, that's a simplified
					// case that doesn't require copying:
					if (!*result)     // Since it's an empty string in their deref buffer,
						result = "";  // ensure it's a non-volatile address instead (read-only mem is okay for expression results).
					else
					{
						// If we don't have have any more function calls pending, we can skip the following step since
						// this deref buffer will not be overwritten during the period we need it.
						for (j = i + 1; j < postfix_count; ++j)
							if (postfix[j]->symbol == SYM_FUNC)
							{
								make_result_persistent = true;
								break;
							}
					}
				}
			}

			if (make_result_persistent)
			{
				// v1.0.44.06: EXPR_SMALL_MEM_LIMIT is the means by which _alloca() is used to boost performance a
				// little by avoiding the overhead of malloc+free for small strings.  The limit should be something
				// small enough that the chance that even 10 of them would cause stack overflow is vanishingly small
				// (the program is currently compiled to allow stack to expand anyway).  Even in a worst-case
				// scenario where an expression is composed entirely of functions and they all need to use this
				// limit of stack space, there's a practical limit on how many functions you can call in an
				// expression due to MAX_TOKENS (probably around MAX_TOKENS / 3).
				#define EXPR_SMALL_MEM_LIMIT 4097
				#define EXPR_ALLOCA_LIMIT 40000  // v1.0.45: Just as an extra precaution against stack stress in extreme/theoretical cases.

				result_size = strlen(result) + 1;
				// Must cast to int to avoid loss of negative values:
				if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
				{
					// Point result to its new, more persistent location:
					result = (char *)memcpy(target, result, result_size); // Benches slightly faster than strcpy().
					target += result_size; // Point it to the location where the next string would be written.
				}
				else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
				{
					result = (char *)memcpy(_alloca(result_size), result, result_size); // Benches slightly faster than strcpy().
					alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
				}
				else // Need to create some new persistent memory for our temporary use.
				{
					// In real-world scripts the need for additonal memory allocation should be quite
					// rare because it requires a combination of worst-case situations:
					// - Called-function's return value is in their new deref buf (rare because return
					//   values are more often literal numbers, true/false, or variables).
					// - We still have more functions to call here (which is somewhat atypical).
					// - There's insufficient room at the end of the deref buf to store the return value
					//   (unusual because the deref buf expands in block-increments, and also because
					//   return values are usually small, such as numbers).
					if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
						|| !(mem[mem_count] = (char *)malloc(result_size)))
					{
						LineError(ERR_OUTOFMEM ERR_ABORT, FAIL, func.mName);
						goto abort;
					}
					// Point result to its new, more persistent location:
					result = (char *)memcpy(mem[mem_count], result, result_size); // Benches slightly faster than strcpy().
					++mem_count; // Must be done last.
				}
			}

			if (!func.mIsBuiltIn)
			{
				// Free the memory of all the just-completed function's local variables.  This is done in
				// both of the following cases:
				// 1) There are other instances of this function beneath us on the call-stack: Must free
				//    the memory to prevent a memory leak for any variable that existed prior to the call
				//    we just did.  Although any local variables newly created as a result of our call
				//    technically don't need to be freed, they are freed for simplicity of code and also
				//    because not doing so might result in side-effects for instances of this function that
				//    lie beneath ours that would expect such nonexistent variables to have blank contents
				//    when *they* create it.
				// 2) No other instances of this function exist on the call stack: The memory is freed and
				//    the contents made blank for these reasons:
				//    a) Prevents locals from all being static in duration, and users coming to rely on that,
				//       since in the future local variables might be implemented using a non-persistent method
				//       such as hashing (rather than maintaining a permanent list of Var*'s for each function).
				//    b) To conserve memory between calls (in case the function's locals use a lot of memory).
				//    c) To yield results consistent with when the same function is called while other instances
				//       of itself exist on the call stack.  In other words, it would be inconsistent to make
				//       all variables blank for case #1 above but not do it here in case #2.
				goto skip_abort_udf;
abort_udf:
// v1.0.40.07: The labels here were added so that an abort of a call to a non-built-in function will collapse
// cleanly.  A goto/label is used rather than some other method because:
// 1) It performs better (the whole reason for having this function be so monolithic is performance).
// 2) Constrained by the performance goal above, it's more maintainable and less code size to do it this way.
				aResult = FAIL;
				early_return = true;
skip_abort_udf:
				for (j = 0; j < func.mVarCount; ++j)
					func.mVar[j]->Free(VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC, true); // Pass "true" to exclude aliases, since their targets should not be freed (they don't belong to this function).
				for (j = 0; j < func.mLazyVarCount; ++j)
					func.mLazyVar[j]->Free(VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC, true);

				// The following used to be the contents of RestoreFunctionVars() [it does too little to be kept
				// as separate a function.]
				// The following call to RestoreFunctionVars() relies on the fact that Free() was already called above.
				// The previous call to BackupFunctionVars() has ensured that none of the variables Free()'d above
				// were ALLOC_SIMPLE, because that would be a memory leak since there's no way to free that type.
				if (var_backup) // This is the indicator that a backup was made, so a restore is also needed.
				{
					for (j = 0; j < var_backup_count; ++j)
						var_backup[j].mVar->Restore(var_backup[j]); // This also frees any existing contents of the variable prior to restoring the original contents from backup.  But it avoids restoring statics.
					free(var_backup);
					var_backup = NULL; // Reset this because it's an indicator of whether the next function call in this expression (if any) will have a backup.
				}

				if (udf_is_final_action) // v1.0.45: An earlier stage has already taken care of this expression's result.
					goto normal_end_skip_output_var; // Nothing more to do because it has even taken care of output_var already.
				// The callers of this function know that the value of aResult (which already contains the reason
				// for early exit) should be considered valid/meaningful only if result_to_return is NULL.
				if (early_return) // aResult has already been set above for our caller.
				{
					result_to_return = NULL; // Use NULL to inform our caller that this thread is finished (whether through normal means such as Exit or a critical error).
					goto normal_end;
				}
			} // if (!func.mIsBuiltIn)

			// Convert this_token's symbol only as the final step in case anything above ever uses its old
			// union member.  Mark it as generic, not string, so that any operator of function call that uses
			// this result is free to reinterpret it as an integer or float:
			this_token.symbol = SYM_OPERAND;
			this_token.marker = result;
			goto push_this_token;
		} // if (this_token.symbol == SYM_FUNC)

		// Since the above didn't "goto", this token must be a unary or binary operator.
		// Get the first operand for this operator (for non-unary operators, this is the right-side operand):
		if (!stack_count) // Prevent stack underflow.  An expression such as -*3 causes this.
			goto abnormal_end;
		ExprTokenType &right = *STACK_POP;
		// Below uses IS_OPERAND rather than checking for only SYM_OPERAND because the stack can contain
		// both generic and specific operands.  Specific operands were evaluated by a previous iteration
		// of this section.  Generic ones were pushed as-is onto the stack by a previous iteration.
		if (!IS_OPERAND(right.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			goto abnormal_end;
		// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
		switch(right.symbol)
		{
		case SYM_VAR:
			right_contents = right.var->Contents();
			right_is_number = IsPureNumeric(right_contents, true, false, true);
			break;
		case SYM_OPERAND:
			right_contents = right.marker;
			right_is_number = IsPureNumeric(right_contents, true, false, true);
			break;
		case SYM_STRING:
			right_contents = right.marker;
			right_is_number = PURE_NOT_NUMERIC; // Explicitly-marked strings are not numeric, which allows numeric strings to be compared as strings rather than as numbers.
		default: // INTEGER or FLOAT
			// right_contents is left uninitialized for performance and to catch bugs.
			right_is_number = right.symbol;
		}

		// IF THIS IS A UNARY OPERATOR, we now have the single operand needed to perform the operation.
		// The cases in the switch() below are all unary operators.  The other operators are handled
		// in the switch()'s default section:
		switch (this_token.symbol)
		{
		case SYM_AND: // These are now unary operators because short-circuit has made them so.  If the AND/OR
		case SYM_OR:  // had short-circuited, we would never be here, so this is the right branch of a non-short-circuit AND/OR.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = (right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents)) != 0;
			else if (right_is_number == PURE_FLOAT)
				this_token.value_int64 = (right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents)) != 0.0;
			else // This is either a non-numeric string or a numeric raw literal string such as "123".
				// All non-numeric strings are considered TRUE here.  In addition, any raw literal string,
				// even "0", is considered to be TRUE.  This relies on the fact that right.symbol will be
				// SYM_OPERAND/generic (and thus handled higher above) for all pure-numeric strings except
				// explicit raw literal strings.  Thus, if something like !"0" ever appears in an expression,
				// it evaluates to !true.  EXCEPTION: Because "if x" evaluates to false when X is blank,
				// it seems best to have "if !x" evaluate to TRUE.
				this_token.value_int64 = *right_contents != '\0';
			this_token.symbol = SYM_INTEGER; // Result of AND or OR is always a boolean integer (one or zero).
			break;

		case SYM_NEGATIVE:  // Unary-minus.
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = -(right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents));
			else if (right_is_number == PURE_FLOAT)
				// Overwrite this_token's union with a float. No need to have the overhead of ATOF() since it can't be hex.
				this_token.value_double = -(right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents));
			else // String.
			{
				// Seems best to consider the application of unary minus to a string, even a quoted string
				// literal such as "15", to be a failure.  UPDATE: For v1.0.25.06, invalid operations like
				// this instead treat the operand as an empty string.  This avoids aborting a long, complex
				// expression entirely just because on of its operands is invalid.  However, the net effect
				// in most cases might be the same, since the empty string is a non-numeric result and thus
				// will cause any operator it is involved with to treat its other operand as a string too.
				// And the result of a math operation on two strings is typically an empty string.
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number; // Convert generic SYM_OPERAND into a specific type: float or int.
			break;

		// Both nots are equivalent at this stage because precedence was already acted upon by infix-to-postfix:
		case SYM_LOWNOT:  // The operator-word "not".
		case SYM_HIGHNOT: // The symbol !
			if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = !(right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents));
			else if (right_is_number == PURE_FLOAT) // Convert to float, not int, so that a number between 0.0001 and 0.9999 is considered "true".
				// Using ! vs. comparing explicitly to 0.0 might generate faster code, and K&R implies it's okay:
				this_token.value_int64 = !(right.symbol == SYM_FLOAT ? right.value_double : atof(right_contents));
			else // This is either a non-numeric string or a numeric raw literal string such as "123".
				// All non-numeric strings are considered TRUE here.  In addition, any raw literal string,
				// even "0", is considered to be TRUE.  This relies on the fact that right.symbol will be
				// SYM_OPERAND/generic (and thus handled higher above) for all pure-numeric strings except
				// explicit raw literal strings.  Thus, if something like !"0" ever appears in an expression,
				// it evaluates to !true.  EXCEPTION: Because "if x" evaluates to false when X is blank,
				// it seems best to have "if !x" evaluate to TRUE.
				this_token.value_int64 = !*right_contents; // i.e. result is false except for empty string because !"string" is false.
			this_token.symbol = SYM_INTEGER; // Result of above is always a boolean integer (one or zero).
			break;

		case SYM_BITNOT: // The tilde (~) operator.
		case SYM_DEREF:  // Dereference an address.
			if (right_is_number == PURE_INTEGER) // But in this case, it can be hex, so use ATOI64().
				right_int64 = right.symbol == SYM_INTEGER ? right.value_int64 : ATOI64(right_contents);
			else if (right_is_number == PURE_FLOAT)
				// No need to have the overhead of ATOI64() since PURE_FLOAT can't be hex:
				right_int64 = right.symbol == SYM_FLOAT ? (__int64)right.value_double : _atoi64(right_contents);
			else // String.  Seems best to consider the application of unary minus to a string, even a quoted string literal such as "15", to be a failure.
			{
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			if (this_token.symbol == SYM_DEREF)
			{
				// Reasons for resolving *Var to a number rather than a single-char string:
				// 1) More consistent with future uses of * that might operate on the address of 2-byte,
				//    4-byte, and 8-byte targets.
				// 2) Performs better in things like ExtractInteger() that would otherwise have to call Asc().
				// 3) Converting it to a one-char string would add no value beyond convenience because script
				//    could do "if (*var = 65)" if it's concerned with avoiding a Chr() call for performance
				//    reasons.  Also, it seems somewhat rare that a script will access a string's characters
				//    one-by-one via the * method because that a parsing loop can already do that more easily.
				// 4) Reduces code size and improves performance (however, the single-char string method would
				//    use _alloca(2) to get some temporary memory, so it wouldn't be too bad in performance).
				//
				// The following does a basic bounds check to prevent crashes due to dereferencing addresses
				// that are obviously bad.  In terms of percentage impact on performance, this seems quite
				// justified.  In the future, could also put a __try/__except block around this (like DllCall
				// uses) to prevent buggy scripts from crashing.  In addition to ruling out the dereferencing of
				// a NULL address, the >255 check also rules out common-bug addresses (I don't think addresses
				// this low can realistically never be legitimate, but it would be nice to get confirmation).
				// For simplicity and due to rarity, a zero is yielded in such cases rather than an empty string.
				// If address is valid, dereference it to extract one unsigned character, just like Asc().
				this_token.value_int64 = (right_int64 < 256 || right_int64 > 0xFFFFFFFF) ? 0 : *(UCHAR *)right_int64;
			}
			else // SYM_BITNOT
			{
				// Note that it is not legal to perform ~, &, |, or ^ on doubles.  Because of this, and also to
				// conform to the behavior of the Transform command, any floating point operand is truncated to
				// an integer above.
				if (right_int64 < 0 || right_int64 > UINT_MAX)
					this_token.value_int64 = ~right_int64;
				else // See comments at TRANS_CMD_BITNOT for why it's done this way:
					this_token.value_int64 = (size_t)~(DWORD)right_int64; // Casting this way avoids compiler warning.
			}
			this_token.symbol = SYM_INTEGER; // Must be done only after its old value was used above. v1.0.36.07: Fixed to be SYM_INTEGER vs. right_is_number for SYM_BITNOT.
			break;

		case SYM_ADDRESS: // Take the address of a variable.
			if (right.symbol == SYM_VAR) // SYM_VAR is always a normal variable, never a built-in one, so taking its address should be safe.
			{
				this_token.symbol = SYM_INTEGER;
				this_token.value_int64 = (__int64)right_contents;
			}
			else // Invalid, so make it a localized blank value.
			{
				this_token.symbol = SYM_STRING;
				this_token.marker = "";
			}
			break;

		default: // Non-unary operator.
			// GET THE SECOND (LEFT-SIDE) OPERAND FOR THIS OPERATOR:
			if (!stack_count) // Prevent stack underflow.
				goto abnormal_end;
			ExprTokenType &left = *STACK_POP; // i.e. the right operand always comes off the stack before the left.
			if (!IS_OPERAND(left.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
				goto abnormal_end;
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
			switch(left.symbol)
			{
			case SYM_VAR:
				left_contents = left.var->Contents();
				left_is_number = IsPureNumeric(left_contents, true, false, true);
				break;
			case SYM_OPERAND:
				left_contents = left.marker;
				left_is_number = IsPureNumeric(left_contents, true, false, true);
				break;
			case SYM_STRING:
				left_contents = left.marker;
				left_is_number = PURE_NOT_NUMERIC;
			default:
				// left_contents is left uninitialized for performance and to catch bugs.
				left_is_number = left.symbol;
			}

			if (!right_is_number || !left_is_number || this_token.symbol == SYM_CONCAT)
			{
				// Above check has ensured that at least one of them is a string.  But the other
				// one might be a number such as in 5+10="15", in which 5+10 would be a numerical
				// result being compared to the raw string literal "15".
				switch (right.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: right_string = ITOA64(right.value_int64, right_buf); break;
				case SYM_FLOAT: snprintf(right_buf, sizeof(right_buf), g.FormatFloat, right.value_double); right_string = right_buf; break;
				default: right_string = right_contents; // SYM_STRING/SYM_OPERAND/SYM_VAR, which is already in the right format.
				}

				switch (left.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: left_string = ITOA64(left.value_int64, left_buf); break;
				case SYM_FLOAT: snprintf(left_buf, sizeof(left_buf), g.FormatFloat, left.value_double); left_string = left_buf; break;
				default: left_string = left_contents; // SYM_STRING or SYM_OPERAND, which is already in the right format.
				}
				
				switch(this_token.symbol)
				{
				case SYM_EQUAL:     this_token.value_int64 = !((g.StringCaseSense == SCS_INSENSITIVE)
										? stricmp(left_string, right_string)
										: lstrcmpi(left_string, right_string)); break; // i.e. use the "more correct mode" except when explicitly told to use the fast mode (v1.0.43.03).
				case SYM_EQUALCASE: this_token.value_int64 = !strcmp(left_string, right_string); break; // Case sensitive.
				// The rest all obey g.StringCaseSense since they have no case sensitive counterparts:
				case SYM_NOTEQUAL:  this_token.value_int64 = g_strcmp(left_string, right_string) ? 1 : 0; break;
				case SYM_GT:        this_token.value_int64 = g_strcmp(left_string, right_string) > 0; break;
				case SYM_LT:        this_token.value_int64 = g_strcmp(left_string, right_string) < 0; break;
				case SYM_GTOE:      this_token.value_int64 = g_strcmp(left_string, right_string) > -1; break;
				case SYM_LTOE:      this_token.value_int64 = g_strcmp(left_string, right_string) < 1; break;

				case SYM_CONCAT:
					// Even if the left or right is "", must copy the result to temporary memory, at least
					// when integers and floats had to be converted to temporary strings above.
					// Calling LengthIgnoreBinaryClip() is valid because SYM_VAR's Type() is always VAR_NORMAL.
					// Binary clipboard is ignored because it's documented that except for certain features,
					// binary clipboard variables are seen only up to the first binary zero (mostly to
					// simplify the code).
					right_length = (right.symbol == SYM_VAR) ? right.var->LengthIgnoreBinaryClip() : strlen(right_string);
					left_length = (left.symbol == SYM_VAR) ? left.var->LengthIgnoreBinaryClip() : strlen(left_string);
					result_size = right_length + left_length + 1;
					// The following section is similar to the one for "symbol == SYM_FUNC", so they
					// should be maintained together.
					// Must cast to int to avoid loss of negative values:
					if (result_size <= (int)(aDerefBufSize - (target - aDerefBuf))) // There is room at the end of our deref buf, so use it.
					{
						this_token.marker = target;
						target += result_size;  // Adjust target for potential future use by another concat or functionc call.
					}
					else if (result_size < EXPR_SMALL_MEM_LIMIT && alloca_usage < EXPR_ALLOCA_LIMIT) // See comments at EXPR_SMALL_MEM_LIMIT.
					{
						this_token.marker = (char *)_alloca(result_size);
						alloca_usage += result_size; // This might put alloca_usage over the limit by as much as EXPR_SMALL_MEM_LIMIT, but that is fine because it's more of a guideline than a limit.
					}
					else // Need to create some new persistent memory for our temporary use.
					{
						// See the nearly identical section higher above for comments:
						if (mem_count == MAX_EXPR_MEM_ITEMS // No more slots left (should be nearly impossible).
							|| !(this_token.marker = mem[mem_count] = (char *)malloc(result_size)))
						{
							LineError(ERR_OUTOFMEM ERR_ABORT);
							goto abort;
						}
						++mem_count;
					}
					if (left_length)
						memcpy(this_token.marker, left_string, left_length);  // Not +1 because don't need the zero terminator.
					memcpy(this_token.marker + left_length, right_string, right_length + 1); // +1 to include its zero terminator.

					// For this new concat operator introduced in v1.0.31, it seems best to treat the
					// result as a SYM_STRING if either operand is a SYM_STRING.  That way, when the
					// result of the operation is later used, it will be a real string even if pure numeric,
					// which allows an exact string match to be specified even when the inputs are
					// technically numeric; e.g. the following should be true only if (Var . 33 = "1133") 
					this_token.symbol = (left.symbol == SYM_STRING || right.symbol == SYM_STRING) ? SYM_STRING: SYM_OPERAND;
					goto push_this_token;

				default:
					// Other operators do not support string operands, so the result is an empty string.
					this_token.marker = "";
					this_token.symbol = SYM_STRING;
					goto push_this_token;
				}
				// Since above didn't "goto":
				this_token.symbol = SYM_INTEGER; // Boolean result is treated as an integer.  Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE
				|| (this_token.symbol == SYM_BITAND || this_token.symbol == SYM_BITOR
				|| this_token.symbol == SYM_BITXOR || this_token.symbol == SYM_BITSHIFTLEFT
				|| this_token.symbol == SYM_BITSHIFTRIGHT)) // The bitwise operators convert any floating points to integer prior to the calculation.
			{
				// Because both are integers and the operation isn't division, the result is integer.
				// The result is also an integer for the bitwise operations listed in the if-statement
				// above.  This is because it is not legal to perform ~, &, |, or ^ on doubles, and also
				// because this behavior conforms to that of the Transform command.  Any floating point
				// operands are truncated to integers prior to doing the bitwise operation.

				switch (right.symbol)
				{
				case SYM_INTEGER: right_int64 = right.value_int64; break;
				case SYM_FLOAT: right_int64 = (__int64)right.value_double; break;
				default: right_int64 = ATOI64(right_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers
				// (otherwise an earlier "else if" would have executed instead of this one).
				}

				switch (left.symbol)
				{
				case SYM_INTEGER: left_int64 = left.value_int64; break;
				case SYM_FLOAT: left_int64 = (__int64)left.value_double; break;
				default: left_int64 = ATOI64(left_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers
				// (otherwise an earlier "else if" would have executed instead of this one).
				}

				switch(this_token.symbol)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case SYM_PLUS:     this_token.value_int64 = left_int64 + right_int64; break;
				case SYM_MINUS:	   this_token.value_int64 = left_int64 - right_int64; break;
				case SYM_TIMES:    this_token.value_int64 = left_int64 * right_int64; break;
				// A look at K&R confirms that relational/comparison operations and logical-AND/OR/NOT
				// always yield a one or a zero rather than arbitrary non-zero values:
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_int64 = left_int64 == right_int64; break;
				case SYM_NOTEQUAL: this_token.value_int64 = left_int64 != right_int64; break;
				case SYM_GT:       this_token.value_int64 = left_int64 > right_int64; break;
				case SYM_LT:       this_token.value_int64 = left_int64 < right_int64; break;
				case SYM_GTOE:     this_token.value_int64 = left_int64 >= right_int64; break;
				case SYM_LTOE:     this_token.value_int64 = left_int64 <= right_int64; break;
				case SYM_BITAND:   this_token.value_int64 = left_int64 & right_int64; break;
				case SYM_BITOR:    this_token.value_int64 = left_int64 | right_int64; break;
				case SYM_BITXOR:   this_token.value_int64 = left_int64 ^ right_int64; break;
				case SYM_BITSHIFTLEFT:  this_token.value_int64 = left_int64 << right_int64; break;
				case SYM_BITSHIFTRIGHT: this_token.value_int64 = left_int64 >> right_int64; break;
				case SYM_FLOORDIVIDE:
					// Since it's integer division, no need for explicit floor() of the result.
					// Also, performance is much higher for integer vs. float division, which is part
					// of the justification for a separate operator.
					if (right_int64 == 0) // Divide by zero produces blank result (perhaps will produce exception if script's ever support exception handlers).
					{
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_token;
					}
					// Otherwise:
					this_token.value_int64 = left_int64 / right_int64;
					break;
				case SYM_POWER:
					// Note: The function pow() in math.h adds about 28 KB of code size (uncompressed)!
					// Even assuming pow() supports negative bases such as (-2)**2, its size is why it's not used.
					// v1.0.44.11: With Laszlo's help, negative integer bases are now supported.
					if (!left_int64 && right_int64 < 0) // In essense, this is divide-by-zero.
					{
						// Return a consistent result rather than something that varies:
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_token;
					}
					// Otherwise, we now have a valid base and exponent and both are integers. Thus, the calculation
					// will always have a defined result.
					if (left_was_negative = (left_int64 < 0))
						left_int64 = -left_int64; // Force a positive due to the limitiations of qmathPow().
					this_token.value_double = qmathPow((double)left_int64, (double)right_int64);
					if (left_was_negative && right_int64 % 2) // Negative base and odd exponent (not zero or even).
						this_token.value_double = -this_token.value_double;
					if (right_int64 < 0)
						this_token.symbol = SYM_FLOAT;  // Due to negative exponent, override to float like TRANS_CMD_POW.
					else
						this_token.value_int64 = (__int64)this_token.value_double;
					break;
				}
				if (this_token.symbol != SYM_FLOAT)  // It wasn't overridden by SYM_POWER.
					this_token.symbol = SYM_INTEGER; // Must be done only after the switch() above.
			}

			else // Since one or both operands are floating point (or this is the division of two integers), the result will be floating point.
			{
				// For these two, use ATOF vs. atof so that if one of them is an integer to be converted
				// to a float for the purpose of this calculation, hex will be supported:
				switch (right.symbol)
				{
				case SYM_INTEGER: right_double = (double)right.value_int64; break;
				case SYM_FLOAT: right_double = right.value_double; break;
				default: right_double = ATOF(right_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch (left.symbol)
				{
				case SYM_INTEGER: left_double = (double)left.value_int64; break;
				case SYM_FLOAT: left_double = left.value_double; break;
				default: left_double = ATOF(left_contents); // SYM_OPERAND or SYM_VAR
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch(this_token.symbol)
				{
				case SYM_PLUS:     this_token.value_double = left_double + right_double; break;
				case SYM_MINUS:	   this_token.value_double = left_double - right_double; break;
				case SYM_TIMES:    this_token.value_double = left_double * right_double; break;
				case SYM_DIVIDE:
				case SYM_FLOORDIVIDE:
					if (right_double == 0.0) // Divide by zero produces blank result (perhaps will produce exception if script's ever support exception handlers).
					{
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_token;
					}
					// Otherwise:
					this_token.value_double = left_double / right_double;
					if (this_token.symbol == SYM_FLOORDIVIDE) // Like Python, the result is floor()'d, moving to the nearest integer to the left on the number line.
						this_token.value_double = qmathFloor(this_token.value_double); // Result is always a double when at least one of the inputs was a double.
					break;
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_double = left_double == right_double; break;
				case SYM_NOTEQUAL: this_token.value_double = left_double != right_double; break;
				case SYM_GT:       this_token.value_double = left_double > right_double; break;
				case SYM_LT:       this_token.value_double = left_double < right_double; break;
				case SYM_GTOE:     this_token.value_double = left_double >= right_double; break;
				case SYM_LTOE:     this_token.value_double = left_double <= right_double; break;
				case SYM_POWER:
					// v1.0.44.11: With Laszlo's help, negative bases are now supported as long as the exponent is not fractional.
					// See the other SYM_POWER higher above for more details about below.
					left_was_negative = (left_double < 0);
					if (left_double == 0.0 && right_double < 0  // In essense, this is divide-by-zero.
						|| left_was_negative && qmathFmod(right_double, 1.0) != 0.0) // Negative base, but exponent isn't close enough to being an integer: unsupported (to simplify code).
					{
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_token;
					}
					// Otherwise:
					if (left_was_negative)
						left_double = -left_double; // Force a positive due to the limitiations of qmathPow().
					this_token.value_double = qmathPow(left_double, right_double);
					if (left_was_negative && qmathFabs(qmathFmod(right_double, 2.0)) == 1.0) // Negative base and exactly-odd exponent (otherwise, it can only be zero or even because if not it would have returned higher above).
						this_token.value_double = -this_token.value_double;
					break;
				}
				this_token.symbol = SYM_FLOAT; // Must be done only after the switch() above.
			} // Result is floating point.
		} // switch() operator type

push_this_token:
		if (!this_token.circuit_token) // It's not capable of short-circuit.
			STACK_PUSH(this_token);    // Push the result onto the stack for use as an operand by a future operator.
		else // This is the final result of a AND or OR's left branch.  Apply short-circuit boolean method to it.
		{
			// Cast this left-branch result to true/false, then determine whether it should cause its
			// parent AND/OR to short-circuit.

			// If its a function result or raw numeric literal such as "if (123 or false)", its type might
			// still be SYM_OPERAND, so resolve that to distinguish between the any SYM_STRING "0"
			// (considered "true") and something that is allowed to be the number zero (which is
			// considered "false").  In other words, the only literal string (or operand made a
			// SYM_STRING via a previous operation) that is considered "false" is the empty string
			// (i.e. "0" doesn't qualify but 0 does):
			switch(this_token.symbol)
			{
			case SYM_VAR:
				// "right" vs. "left" is used even though this is technically the left branch because
				// right is used more often (for unary operators) and sometimes the compiler generates
				// faster code for the most frequently accessed variables.
				right_contents = this_token.var->Contents();
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_OPERAND:
				right_contents = this_token.marker;
				right_is_number = IsPureNumeric(right_contents, true, false, true);
				break;
			case SYM_STRING:
				right_contents = this_token.marker;
				right_is_number = PURE_NOT_NUMERIC;
			default:
				// right_contents is left uninitialized for performance and to catch bugs.
				right_is_number = this_token.symbol;
			}

			switch (right_is_number)
			{
			case PURE_INTEGER: // Probably the most common, e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
				// Force it to be purely 1 or 0 if it isn't already.
				left_branch_is_true = (this_token.symbol == SYM_INTEGER ? this_token.value_int64
					: ATOI64(right_contents)) != 0;
				break;
			case PURE_FLOAT: // Convert to float, not int, so that a number between 0.0001 and 0.9999 is is considered "true".
				left_branch_is_true = (this_token.symbol == SYM_FLOAT ? this_token.value_double
					: atof(right_contents)) != 0.0;
				break;
			default:  // string.
				// Since "if x" evaluates to false when x is blank, it seems best to also have blank
				// strings resolve to false when used in more complex ways. In other words "if x or y"
				// should be false if both x and y are blank.  Logical-not also follows this convention.
				left_branch_is_true = (*right_contents != '\0');
			}

			// The following loop exists to support cascading short-circuiting such as the following example:
			// 2>3 and 2>3 and 2>3
			// In postfix notation, the above looks like:
			// 2 3 > 2 3 > and 2 3 > and
			// When the first '>' operator is evaluated to false, it sees that its parent is an AND and
			// thus it short-circuits, discarding everything between the first '>' and the "and".
			// But since the first and's parent is the second "and", that false result just produced is now
			// the left branch of the second "and", so the loop conducts a second iteration to discard
			// everything between the first "and" and the second.  By contrast, if the second "and" were
			// an "or", the second iteration would never occur because the loop's condition would be false
			// on the second iteration, which would then cause the first and's false value to be discarded
			// (due to the loop ending without having PUSHed) because solely the right side of the "or" should
			// determine the final result of the "or".
			for (circuit_token = this_token.circuit_token
				; left_branch_is_true == (circuit_token->symbol == SYM_OR);) // If true, this AND/OR causes a short-circuit
			{
				// Discard the entire right branch of this AND/OR:
				for (++i; postfix[i] != circuit_token; ++i); // Should always be found, so no need to check postfix_count.
				// Above loop is self-contained.
				if (   !(circuit_token = postfix[i]->circuit_token)   ) // This value is also used by our loop's condition.
				{
					// No more cascading is needed because this AND/OR isn't the left branch of another.
					// This will be the final result of this AND/OR because it's right branch was discarded
					// above without having been evaluated nor any of its functions called.  It's safe to use
					// this_token vs. postfix[i] below, for performance, because the value in its circuit_token
					// member no longer matters:
					this_token.symbol = SYM_INTEGER;
					this_token.value_int64 = left_branch_is_true; // Assign a pure 1 (for SYM_OR) or 0 (for SYM_AND).
					STACK_PUSH(this_token);
					break; // Now the outer loop's ++i will discard this AND/OR token itself and continue onward.
				}
				//else there is more cascading to be checked, so continue looping.
			}
			// If the loop ends normally (not via "break"), postfix[i] is now the left branch of an
			// AND/OR that should not short-circuit.  As a result, this left branch is simply discarded
			// (by means of the outer loop's ++i) because its right branch will be the sole determination
			// of whether this AND/OR is true or false.
		} // Short-circuit (left branch of an AND/OR).
	} // For each item in the postfix array.

	// Although ACT_FUNCTIONCALL was already checked higher above, it's checked again here for maintainability.
	// Specifically, there might be ways the above didn't return if ACT_FUNCTIONCALL, such as when somehow there
	// was more than one token on the stack even for the final function call, or maybe other unforseen ways.
	// It seems best to avoid any chance of looking at the result since it might be invalid due to the above
	// having taken shortcuts (since it knew the result wouldn't be needed).
	if (mActionType == ACT_FUNCTIONCALL) // A line consisting only of a function call (possibly with nested function calls): the end result doesn't matter, even if it's a failure.
		goto normal_end; // And leave result_to_return at its default of "".

	if (stack_count != 1) // Stack should have only one item left on it: the result. If not, it's a syntax error.
		goto abnormal_end; // This with these examples: 1) (); 2) x y; 3) (x + y) (x + z); etc.

	ExprTokenType &result_token = *stack[0];  // For performance and convenience.

	// Store the result of the expression in the deref buffer for the caller.  It is stored in the current
	// format in effect via SetFormat because:
	// 1) The := operator then doesn't have to convert to int/double then back to string to put the right format into effect.
	// 2) It might add a little bit of flexibility in places parameters where floating point values are expected
	//    (i.e. it allows a way to do automatic rounding), without giving up too much.  Changing floating point
	//    precision from the default of 6 decimal places is rare anyway, so as long as this behavior is documented,
	//    it seems okay for the moment.
	switch (result_token.symbol)
	{
	// The numeric cases below will always fit into our deref buffer because an earlier stage has already ensured
	// that the buffer is large enough to hold at least one number.  But a string/generic might not fit if it's
	// a concatenation and/or a large string returned from a called function.
	case SYM_INTEGER:
		// To reduce code size, the following isn't done (it's unlikely to help performance anyway if you think
		// about what steps it actually takes):
		//if (output_var)
		//{
		//	output_var->Assign(result_token.value_int64);
		//	goto normal_end_skip_output_var;
		//}
		//// Otherwise:
		ITOA64(result_token.value_int64, aTarget); // Store in hex or decimal format, as appropriate.
		break;
	case SYM_FLOAT:
		// To reduce code size, the following isn't done (it's unlikely to help performance anyway if you think
		// about what steps it actually takes):
		//if (output_var)
		//{
		//	output_var->Assign(result_token.value_double);
		//	goto normal_end_skip_output_var;
		//}
		//// Otherwise:
		// In case of float formats that are too long to be supported, use snprint() to restrict the length.
		snprintf(aTarget, MAX_FORMATTED_NUMBER_LENGTH + 1, g.FormatFloat, result_token.value_double); // %f probably defaults to %0.6f.  %f can handle doubles in MSVC++.
		break;

	case SYM_STRING:
	case SYM_OPERAND:
	case SYM_VAR: // SYM_VAR is somewhat unusual at this late a stage.
		// At this stage, we know the result has to go into our deref buffer because if a way existed to
		// avoid that, we would already have goto/returned higher above.  Also, at this stage,
		// the pending result can exist in one several places:
		// 1) Our deref buf (due to being a single-deref, a function's return value that was copied to the
		//    end of our buf because there was enough room, etc.)
		// 2) In a called function's deref buffer, namely sDerefBuf, which will be deleted by our caller
		//    shortly after we return to it.
		// 3) In an area of memory we malloc'd for lack of any better place to put it.
		char *result;
		if (result_token.symbol == SYM_VAR)
		{
			result = result_token.var->Contents();
            result_size = result_token.var->LengthIgnoreBinaryClip() + 1; // Ignore binary clipboard because it's documented that except for certain features, binary clipboard variables are seen only up to the first binary zero (mostly to simplify the code).
		}
		else
		{
			result = result_token.marker;
			result_size = strlen(result) + 1;
		}
		if (output_var)
		{
			// v1.0.45: Take a shortcut, which avoids one memcpy (into the deref buffer).  In some cases, this
			// also saves from having to expand the deref buffer.
			output_var->Assign(result, (VarSizeType)result_size - 1);
			goto normal_end_skip_output_var;
		}
		// Otherwise:
		// If result is the empty string or a number, it should always fit because the size estimation
		// phase has ensured that capacity_of_our_buf_portion is large enough to hold those:
		if (result_size > capacity_of_our_buf_portion)
		{
			// Do a simple expansion of our deref buffer to handle the fact that our actual result is bigger
			// than the size estimator could have calculated (due to a concatenation or a large string returned
			// from a called function).  This performs poorly but seems justified by the fact that it is
			// typically needed only in extreme cases.
			size_t new_buf_size = aDerefBufSize + result_size - capacity_of_our_buf_portion;

			// malloc() and free() are used instead of realloc() because in many cases, the overhead of
			// realloc()'s internal memcpy(entire contents) can be avoided because only part or
			// none of the contents needs to be copied (realloc's ability to do an in-place resize might
			// be unlikely for anything other than small blocks; see compiler's realloc.c):
			char *new_buf = (char *)malloc(new_buf_size);
			if (!new_buf)
			{
				LineError(ERR_OUTOFMEM ERR_ABORT);
				goto abort;
			}
			if (new_buf_size > LARGE_DEREF_BUF_SIZE)
				++sLargeDerefBufs;

			// Copy only that portion of the old buffer that is in front of our portion of the buffer
			// because we no longer need our portion (except for result.marker if it happens to be
			// in the old buffer, but that is handled after this):
			size_t aTarget_offset = aTarget - aDerefBuf;
			if (aTarget_offset) // aDerefBuf has contents that must be preserved.
				memcpy(new_buf, aDerefBuf, aTarget_offset); // This will also copy the empty string if the buffer first and only character is that.
			aTarget = new_buf + aTarget_offset;
			// NOTE: result_token.marker might be at the end of our deref buffer and thus be larger than
			// capacity_of_our_buf_portion because other arg(s) exist in this line after ours that will be
			// using a larger total portion of the buffer than ours.  Thus, the following must be done prior
			// to free(), but memcpy() vs. memmove() is safe in any case:
			memcpy(aTarget, result, result_size); // Copy from old location to the newly allocated one.

			free(aDerefBuf); // Free our original buffer since it's contents are no longer needed.
			if (aDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;

			// Now that the buffer has been enlarged, need to adjust any other pointers that pointed into
			// the old buffer:
			char *aDerefBuf_end = aDerefBuf + aDerefBufSize; // Point it to the character after the end of the old buf.
			for (i = 0; i < aArgIndex; ++i) // Adjust each item beneath ours (if any). Our own is not adjusted because we'll be returning the right address to our caller.
				if (aArgDeref[i] >= aDerefBuf && aArgDeref[i] < aDerefBuf_end)
					aArgDeref[i] = new_buf + (aArgDeref[i] - aDerefBuf); // Set for our caller.
			// The following isn't done because target isn't used anymore at this late a stage:
			//target = new_buf + (target - aDerefBuf);
			aDerefBuf = new_buf; // Must be the last step, since the old address is used above.  Set for our caller.
			aDerefBufSize = new_buf_size; // Set for our caller.
		}
		else // Deref buf is already large enough to fit the string.
			if (aTarget != result) // Currently, might be always true.
				memmove(aTarget, result, result_size); // memmove() vs. memcpy() in this case, since source and dest might overlap.
		result_to_return = aTarget;
		aTarget += result_size;
		goto normal_end;

	default: // Result contains a non-operand symbol such as an operator.
		goto abnormal_end;
	}

	// Since above didn't "goto", this is SYM_FLOAT/SYM_INTEGER.  Calculate the length and use it to adjust
	// aTarget for use by our caller:
	result_to_return = aTarget;
	aTarget += strlen(aTarget) + 1;  // +1 because that's what callers want; i.e. the position after the terminator.
	goto normal_end;

abort:
	// The callers of this function know that the value of aResult (which contains the reason
	// for early exit) should be considered valid/meaningful only if result_to_return is NULL.
	result_to_return = NULL; // Use NULL to inform our caller that this entire thread is to be terminated.
	aResult = FAIL; // Indicate reason to caller.

abnormal_end: // Currently the same as normal_end; it's separate to improve readability.  When this happens, result_to_return is typically "" (unless the caller overrode that default).
normal_end:
	// v1.0.45: ACT_ASSIGNEXPR relies on us to set the output_var (i.e. whenever it's ARG1's is_expression==true).
	// Our taking charge of output_var allows certain performance optimizations in other parts of this function,
	// such as avoiding excess memcpy's and malloc's during intermediate stages.
	if (output_var && result_to_return) // i.e. don't assign if NULL to preserve backward compatibility with scripts that rely on the old value being changed in cases where an expression fails (unlikely).
		output_var->Assign(result_to_return);

normal_end_skip_output_var:
	for (i = 0; i < mem_count; ++i) // Free any temporary memory blocks that were used.
		free(mem[i]);

	return result_to_return;
}



ResultType Line::ExpandArgs(VarSizeType aSpaceNeeded, Var *aArgVar[])
// Caller should either provide both or omit both of the parameters.  If provided, it means
// caller already called GetExpandedArgSize for us.
// Returns OK, FAIL, or EARLY_EXIT.  EARLY_EXIT occurs when a function-call inside an expression
// used the EXIT command to terminate the thread.
{
	// The counterparts of sArgDeref and sArgVar kept on our stack to protect them from recursion caused by
	// the calling of functions in the script:
	char *arg_deref[MAX_ARGS];
	Var *arg_var[MAX_ARGS];
	int i;

	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former often does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how much space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	// First pass. It takes into account the same things as 2nd pass.
	size_t space_needed;
	if (aSpaceNeeded == VARSIZE_ERROR)
	{
		space_needed = GetExpandedArgSize(true, arg_var);
		if (space_needed == VARSIZE_ERROR)
			return FAIL;  // It will have already displayed the error.
	}
	else // Caller already determined it.
	{
		space_needed = aSpaceNeeded;
		for (i = 0; i < mArgc; ++i) // Copying only the actual/used elements is probably faster than using memcpy to copy both entire arrays.
			arg_var[i] = aArgVar[i]; // Init to values determined by caller, which helps performance if any of the args are dynamic variables.
	}

	if (space_needed > g_MaxVarCapacity)
		// Dereferencing the variables in this line's parameters would exceed the allowed size of the temp buffer:
		return LineError(ERR_MEM_LIMIT_REACHED);

	// Only allocate the buf at the last possible moment,
	// when it's sure the buffer will be used (improves performance when only a short
	// script with no derefs is being run):
	if (space_needed > sDerefBufSize)
	{
		size_t increments_needed = space_needed / DEREF_BUF_EXPAND_INCREMENT;
		if (space_needed % DEREF_BUF_EXPAND_INCREMENT)  // Need one more if above division truncated it.
			++increments_needed;
		size_t new_buf_size = increments_needed * DEREF_BUF_EXPAND_INCREMENT;
		if (sDerefBuf)
		{
			// Do a free() and malloc(), which should be far more efficient than realloc(), especially if
			// there is a large amount of memory involved here (realloc's ability to do an in-place resize
			// might be unlikely for anything other than small blocks; see compiler's realloc.c):
			free(sDerefBuf);
			if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;
		}
		if (   !(sDerefBuf = (char *)malloc(new_buf_size))   )
		{
			// Error msg was formerly: "Ran out of memory while attempting to dereference this line's parameters."
			sDerefBufSize = 0;  // Reset so that it can make another attempt, possibly smaller, next time.
			return LineError(ERR_OUTOFMEM ERR_ABORT); // Short msg since so rare.
		}
		sDerefBufSize = new_buf_size;
		if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
			++sLargeDerefBufs;
	}

	// Always init our_buf_marker even if zero iterations, because we want to enforce
	// the fact that its prior contents become invalid once we're called.
	// It's also necessary due to the fact that all the old memory is discarded by
	// the above if more space was needed to accommodate this line.
	char *our_buf_marker = sDerefBuf;  // Prior contents of buffer will be overwritten in any case.

	// From this point forward, must not refer to sDerefBuf as our buffer since it might have been
	// given a new memory area by an expression's function-call within this line.  In other words,
	// our_buf_marker is our recursion layer's buffer, but not necessarily sDerefBuf.  To enforce
	// that, and keep responsibility here rather than in ExpandExpression(), set sDerefBuf to NULL
	// so that the zero or more calls to ExpandExpression() made in the loop below, which in turn
	// will result in zero or more invocations of script-functions, will allocate and use a single
	// new deref buffer if any of them need it.
	// Note that it is not possible for a new quasi-thread to directly interrupt ExpandArgs(),
	// since ExpandArgs() never calls MsgSleep().  Therefore, each ExpandArgs() layer on the call
	// stack is safe from interrupting threads overwriting its deref buffer.  It's true that a call
	// to a script function will usually result in MsgSleep(), and thus allow interruptions, but those
	// interruptions would hit some other deref buffer, not that of our layer.
	char *our_deref_buf = sDerefBuf; // For detecting whether ExpandExpression() caused a new buffer to be created.
	size_t our_deref_buf_size = sDerefBufSize;
	sDerefBuf = NULL;
	sDerefBufSize = 0;

	ResultType result, result_to_return = OK;  // Set default return value.
	Var *the_only_var_of_this_arg;

	if (!mArgc)            // v1.0.45: Required by some commands that can have zero parameters (such as Random and
		sArgVar[0] = NULL; // PixelSearch), even if it's just to allow their output-var(s) to be omitted.  This allows OUTPUT_VAR to be used without any need to check mArgC.
	else
	{
		for (i = 0; i < mArgc; ++i) // Second pass.  For each arg:
		{
			ArgStruct &this_arg = mArg[i]; // For performance and convenience.

			// Load-time routines have already ensured that an arg can be an expression only if
			// it's not an input or output var.
			if (this_arg.is_expression)
			{
				// v1.0.45:
				// Make ARGVAR1 (OUTPUT_VAR) temporarily valid (the entire array is made valid only later, near the
				// bottom of this function).  This helps the performance of ACT_ASSIGNEXPR by avoiding the need
				// resolve a dynamic output variable like "Array%i% := (Expr)" twice: once in GetExpandedArgSize
				// and again in ExpandExpression()).
				*sArgVar = *arg_var; // Shouldn't need to be backed up or restored because no one beneath us on the call stack should be using it; only things that go on top of us might overwrite it, so ExpandExpr() must be sure to copy this out before it launches any script-functions.
				// In addition to producing its return value, ExpandExpression() will alter our_buf_marker
				// to point to the place in our_deref_buf where the next arg should be written.
				// In addition, in some cases it will alter some of the other parameters that are arrays or
				// that are passed by ref.  Finally, it might tempoarily use parts of the buffer beyond
				// what the size estimator provide for it, so we should be sure here that everything in
				// our_deref_buf after our_buf_marker is available to it as temporary memory.
				// Note: It doesn't seem worthwhile to enhance ExpandExpression to give us back a variable
				// for use in arg_var[] (for performance) because only rarely does an expression yield
				// a variable other than some function's local variable (and a local's contents are no
				// longer valid due to having been freed after the call [unless it's static]).
				if (   !(arg_deref[i] = ExpandExpression(i, result, our_buf_marker, our_deref_buf
					, our_deref_buf_size, arg_deref, our_deref_buf_size - space_needed))   )
				{
					// A script-function-call inside the expression returned EARLY_EXIT or FAIL.  Report "result"
					// to our caller (otherwise, the contents of "result" should be ignored since they're undefined).
					result_to_return = result;
					goto end;
				}
				continue;
			}

			if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // Don't bother wasting the mem to deref output var.
			{
				// In case its "dereferenced" contents are ever directly examined, set it to be
				// the empty string.  This also allows the ARG to be passed a dummy param, which
				// makes things more convenient and maintainable in other places:
				arg_deref[i] = "";
				continue;
			}

			// arg_var[i] was previously set by GetExpandedArgSize() so that we don't have to determine its
			// value again:
			if (   !(the_only_var_of_this_arg = arg_var[i])   ) // Arg isn't an input var or singled isolated deref.
			{
				#define NO_DEREF (!ArgHasDeref(i + 1))
				if (NO_DEREF)
				{
					arg_deref[i] = this_arg.text;  // Point the dereferenced arg to the arg text itself.
					continue;  // Don't need to use the deref buffer in this case.
				}
			}

			// Check the value of the_only_var_of_this_arg again in case the above changed it:
			if (the_only_var_of_this_arg) // This arg resolves to only a single, naked var.
			{
				switch(ArgMustBeDereferenced(the_only_var_of_this_arg, i, arg_var)) // Yes, it was called by GetExpandedArgSize() too, but a review shows it's difficult to avoid this without being worse than the disease (10/22/2006).
				{
				case CONDITION_FALSE:
					// This arg contains only a single dereference variable, and no
					// other text at all.  So rather than copy the contents into the
					// temp buffer, it's much better for performance (especially for
					// potentially huge variables like %clipboard%) to simply set
					// the pointer to be the variable itself.  However, this can only
					// be done if the var is the clipboard or a non-environment
					// normal var (since zero-length normal vars need to be fetched via
					// GetEnvironmentVariable() when g_NoEnv==false).
					// Update: Changed it so that it will deref the clipboard if it contains only
					// files and no text, so that the files will be transcribed into the deref buffer.
					// This is because the clipboard object needs a memory area into which to write
					// the filespecs it translated:
					arg_deref[i] = the_only_var_of_this_arg->Contents();
					break;
				case CONDITION_TRUE:
					// the_only_var_of_this_arg is either a reserved var or a normal var of that is also
					// an environment var (for which GetEnvironmentVariable() is called for), or is used
					// again in this line as an output variable.  In all these cases, it must
					// be expanded into the buffer rather than accessed directly:
					arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
					our_buf_marker += the_only_var_of_this_arg->Get(our_buf_marker) + 1; // +1 for terminator.
					break;
				default: // FAIL should be the only other possibility.
					result_to_return = FAIL; // ArgMustBeDereferenced() will already have displayed the error.
					goto end;
				}
			}
			else // The arg must be expanded in the normal, lower-performance way.
			{
				arg_deref[i] = our_buf_marker; // Point it to its location in the buffer.
				if (   !(our_buf_marker = ExpandArg(our_buf_marker, i))   ) // Expand the arg into that location.
				{
					result_to_return = FAIL; // ExpandArg() will have already displayed the error.
					goto end;
				}
			}
		} // for each arg.

		// IT'S NOT SAFE to do the following until the above loop FULLY completes because any calls made above to
		// ExpandExpression() might call functions, which in turn might result in a recursive call to ExpandArgs(),
		// which in turn might change the values in the static arrays sArgDeref and sArgVar.
		// Also, only when the loop ends normally is the following needed, since otherwise it's a failure condition.
		// Now that any recursive calls to ExpandArgs() above us on the stack have collapsed back to us, it's
		// safe to set the args of this command for use by our caller, to whom we're about to return.
		for (i = 0; i < mArgc; ++i) // Copying actual/used elements is probably faster than using memcpy to copy both entire arrays.
		{
			sArgDeref[i] = arg_deref[i];
			sArgVar[i] = arg_var[i];
		}
	} // mArgc > 0

	// v1.0.40.02: The following loop was added to avoid the need for the ARGn macros to provide an empty
	// string when mArgc was too small (indicating that the parameter is absent).  This saves quite a bit
	// of code size.  Also, the slight performance loss caused by it is partially made up for by the fact
	// that all the other sections don't need to check mArgc anymore.
	// Benchmarks show that it doesn't help performance to try to tweak this with a pre-check such as
	// "if (mArgc < max_params)":
	int max_params = g_act[mActionType].MaxParams; // Resolve once for performance.
	for (i = mArgc; i < max_params; ++i) // For performance, this only does the actual max args for THIS command, not MAX_ARGS.
		sArgDeref[i] = "";
		// But sArgVar isn't done (since it's more rarely used) except sArgVar[0] = NULL higher above.
		// Therefore, users of sArgVar must check mArgC if they have any doubt how many args are present in
		// the script line (this is now enforced via macros).

	// When the main/large loop above ends normally, it falls into the label below and uses the original/default
	// value of "result_to_return".

end:
	// As of v1.0.31, there can be multiple deref buffers simultaneously if one or more called functions
	// requires a deref buffer of its own (separate from ours).  In addition, if a called function is
	// interrupted by a new thread before it finishes, the interrupting thread will also use the
	// new/separate deref buffer.  To minimize the amount of memory used in such cases cases,
	// each line containing one or more expression with one or more function call (rather than each
	// function call) will get up to one deref buffer of its own (i.e. only if its function body contains
	// commands that actually require a second deref buffer).  This is achieved by saving sDerefBuf's
	// pointer and setting sDerefBuf to NULL, which effectively makes the original deref buffer private
	// until the line that contains the function-calling expressions finishes completely.
	// Description of recursion and usage of multiple deref buffers:
	// 1) ExpandArgs() receives a line with one or more expressions containing one or more function-calls.
	// 2) Worst-case: the function calls create a new sDerefBuf automatically via us having set sDerefBuf to NULL.
	// 3) Even worse, the bodies of those functions call other functions, which ExpandArgs() receives, resulting in
	// a recursive leap back to step #1.
	// So the above shows how any number of new deref buffers can be created.  But that's okay as long as the
	// recursion collapses in an orderly manner (or the program exits, in which case the OS frees all its memory
	// automatically).  This is because prior to returning, each recursion layer properly frees any extra deref
	// buffer it was responsible for creating.  It only has to free at most one such buffer because each layer of
	// ExpandArgs() on the call-stack can never be blamed for creating more than one extra buffer.
	if (our_deref_buf)
	{
		// Must always restore the original buffer, not the keep the new one, because our caller needs
		// the arg_deref addresses, which point into the original buffer.
		if (sDerefBuf)
		{
			free(sDerefBuf);
			if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
				--sLargeDerefBufs;
		}
		sDerefBuf = our_deref_buf;
		sDerefBufSize = our_deref_buf_size;
	}
	//else the original buffer is NULL, so keep any new sDerefBuf that might have been created (should
	// help avg-case performance).

	// For v1.0.31, this is no done right before returning so that any script function calls
	// made by our calls to ExpandExpression() will now be done.  There might still be layers
	// of ExpandArgs() beneath us on the call-stack, which is okay since they will keep the
	// largest of the two available deref bufs (as described above) and thus they should
	// reset the timer below right before they collapse/return.  
	// (Re)set the timer unconditionally so that it starts counting again from time zero.
	// In other words, we only want the timer to fire when the large deref buffer has been
	// unused/idle for a straight 10 seconds.  There is no danger of this timer freeing
	// the deref buffer at a critical moment because:
	// 1) The timer is reset with each call to ExpandArgs (this function);
	// 2) If our ExpandArgs() recursion layer takes a long time to finish, messages
	//    won't be checked and thus the timer can't fire because it relies on the msg loop.
	// 3) If our ExpandArgs() recursion layer launches function-calls in ExpandExpression(),
	//    those calls will call ExpandArgs() recursively and reset the timer if its
	//    buffer (not necessarily the original buffer somewhere on the call-stack) is large
	//    enough.  In light of this, there is a chance that the timer might execute and free
	//    a deref buffer other than the one it was originally intended for.  But in real world
	//    scenarios, that seems rare.  In addition, the consequences seem to be limited to
	//    some slight memory inefficiency.
	// It could be aruged that the timer should only be activated when a hypothetical static
	// var sLayersthat we maintain here indicates that we're the only layer.  However, if that
	// were done and the launch of a script function creates (directly or through thread
	// interruption, indirectly) a large deref buffer, and that thread is waiting for something
	// such as WinWait, that large deref buffer would never get freed.
	#define SET_DEREF_TIMER(aTimeoutValue) g_DerefTimerExists = SetTimer(g_hWnd, TIMER_ID_DEREF, aTimeoutValue, DerefTimeout);
	if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
		SET_DEREF_TIMER(10000) // Reset the timer right before the deref buf is possibly about to become idle.

	return result_to_return;
}

	

VarSizeType Line::GetExpandedArgSize(bool aCalcDerefBufSize, Var *aArgVar[])
// Args that are expressions are only calculated correctly if aCalcDerefBufSize is true,
// which is okay for the moment since the only caller that can have expressions does call
// it that way.
// Returns the size, or VARSIZE_ERROR if there was a problem.
// WARNING: This function can return a size larger than what winds up actually being needed
// (e.g. caused by ScriptGetCursor()), so our callers should be aware that that can happen.
{
	int i;
	VarSizeType space_needed, space;
	DerefType *deref;
	Var *the_only_var_of_this_arg;
	bool include_this_arg;
	ResultType result;

	// Note: the below loop is similar to the one in ExpandArgs(), so the two should be
	// maintained together:
	for (i = 0, space_needed = 0; i < mArgc; ++i) // For each arg:
	{
		ArgStruct &this_arg = mArg[i]; // For performance and convenience.

		// If this_arg.is_expression is true, the space is still calculated as though the
		// expression itself will be inside the arg.  This is done so that an expression
		// such as if(Array%i% = LargeString) can be expanded temporarily into the deref
		// buffer so that it can be evaluated more easily.

		// Accumulate the total of how much space we will need.
		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // These should never be included in the space calculation.
		{
			if (mActionType != ACT_ASSIGN) // PerformAssign() already resolve its output-var, so don't do it again here.
			{
				if (   !(aArgVar[i] = ResolveVarOfArg(i))   ) // v1.0.45: Resolve output variables too, which eliminates a ton of calls to ResolveVarOfArg() in various other functions.  This helps code size more than performance.
					return VARSIZE_ERROR;  // The above will have already displayed the error.
			}
			continue;
		}
		// Otherwise, set default.
		aArgVar[i] = NULL;

		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		the_only_var_of_this_arg = NULL;
		if (this_arg.type == ARG_TYPE_INPUT_VAR) // Previous stage has ensured that arg can't be an expression if it's an input var.
			if (   !(the_only_var_of_this_arg = ResolveVarOfArg(i, false))   )
				return VARSIZE_ERROR;  // The above will have already displayed the error.

		if (!the_only_var_of_this_arg) // It's not an input var.
		{
			if (NO_DEREF)
			{
				// Below relies on the fact that caller has ensure no args are expressions
				// when !aCalcDerefBufSize.
				if (!aCalcDerefBufSize || this_arg.is_expression) // i.e. we want the total size of what the args resolve to.
					space_needed += this_arg.length + 1;  // +1 for the zero terminator.
				// else don't increase space_needed, even by 1 for the zero terminator, because
				// the terminator isn't needed if the arg won't exist in the buffer at all.
				continue;
			}
			// Now we know it has at least one deref.  If the second deref's marker is NULL,
			// the first is the only deref in this arg.  UPDATE: The following will return
			// false for function calls since they are always followed by a set of parentheses
			// (empty or otherwise), thus they will never be seen as isolated by it:
			#define SINGLE_ISOLATED_DEREF (!this_arg.deref[1].marker\
				&& this_arg.deref[0].length == this_arg.length) // and the arg contains no literal text
			if (SINGLE_ISOLATED_DEREF) // This also ensures the deref isn't a function-call.  10/25/2006: It might be possible to avoid the need for detecting SINGLE_ISOLATED_DEREF by transforming them into INPUT_VARs at loadtime.  I almost finished such a mod but the testing and complications with things like ListLines didn't seem worth the tiny benefit.
				the_only_var_of_this_arg = this_arg.deref[0].var;
		}
		if (the_only_var_of_this_arg) // i.e. check it again in case the above block changed the value.
		{
			// This is set for our caller so that it doesn't have to call ResolveVarOfArg() again, which
			// would a performance hit if this variable is dynamically built and thus searched for at runtime:
			aArgVar[i] = the_only_var_of_this_arg; // For now, this is done regardless of whether it must be dereferenced.
			include_this_arg = !aCalcDerefBufSize || this_arg.is_expression;  // i.e. caller wanted its size unconditionally included
			if (!include_this_arg)
			{
				if (   !(result = ArgMustBeDereferenced(the_only_var_of_this_arg, i, aArgVar))   )
					return VARSIZE_ERROR;
				if (result == CONDITION_TRUE) // The size of these types of args is always included.
					include_this_arg = true;
				//else leave it as false
			}
			if (!include_this_arg) // No extra space is needed in the buffer for this arg.
				continue;
			space = the_only_var_of_this_arg->Get() + 1;  // +1 for the zero terminator.
			// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
			// being needed, so our callers should be aware that that can happen.
			if (this_arg.is_expression) // Space is needed for the result of the expression or the expanded expression itself, whichever is greater.
				space_needed += (space > MAX_FORMATTED_NUMBER_LENGTH) ? space : MAX_FORMATTED_NUMBER_LENGTH + 1;
			else
				space_needed += space;
			continue;
		}

		// Otherwise: This arg has more than one deref, or a single deref with some literal text around it.
		space = this_arg.length + 1; // +1 for this arg's zero terminator in the buffer.
		for (deref = this_arg.deref; deref && deref->marker; ++deref)
		{
			// Replace the length of the deref's literal text with the length of its variable's contents:
			space -= deref->length;
			if (!deref->is_function)
			{
				if (this_arg.is_expression)
				{
					// In the case of expressions, size needs to be reserved for the variable's contents only
					// if it will be copied into the deref buffer; namely the following cases:
					// 1) Derefs whose type isn't VAR_NORMAL or that are env. vars (those whose length is zero but whose Get() is of non-zero length)
					// 2) Derefs that are enclosed by the g_DerefChar character (%), which in expressions means that
					//    must be copied into the buffer to support double references such as Array%i%.
					if (*deref->marker == g_DerefChar || deref->var->Type() != VAR_NORMAL // Relies on short-circuit boolean order for the next line.
						|| (!g_NoEnv && !deref->var->Length())) // v1.0.43.08: Added g_NoEnv.
						space += deref->var->Get(); // If an environment var, Get() will yield its length.
					space += 1;
					// Fix for v1.0.35.04: The above now adds a space unconditionally because it is needed
					// by the expression evaluation to provide an empty string (terminator) in the deref 
					// buf for each variable, which prevents something like "x*y*z" from being seen as
					// two asterisks in a row (since y doesn't take up any space).  Although the +1 might
					// not be needed in a few sub-cases of the above, it is safer to do it and doesn't
					// increase the size much anyway.  Note that function-calls do not need this fix because
					// their parentheses and arg list are always in the deref buffer.
					// Above adds 1 for the insertion of an extra space after every single deref.  This space
					// is unnecessary if Get() returns a size of zero to indicate a non-existent environment
					// variable, but that seems harmless).  This is done for parsing reasons described in
					// ExpandExpression().
					// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
					// being needed, so our callers should be aware that that can happen.
				}
				else // Not an expression.
					space += deref->var->Get(); // If an environment var, Get() will yield its length.
			}
			//else it's a function-call's function name, in which case it's length is effectively zero.
			// since the function name never gets copied into the deref buffer during ExpandExpression().
		}
		if (this_arg.is_expression) // Space is needed for the result of the expression or the expanded expression itself, whichever is greater.
			space_needed += (space > MAX_FORMATTED_NUMBER_LENGTH) ? space : MAX_FORMATTED_NUMBER_LENGTH + 1;
		else
			space_needed += space;
	}
	return space_needed;
}



ResultType Line::ArgMustBeDereferenced(Var *aVar, int aArgIndex, Var *aArgVar[]) // 10/22/2006: __forceinline didn't help enough to be worth the added code size of having two instances.
// Shouldn't be called only for args of type ARG_TYPE_OUTPUT_VAR because they never need to be dereferenced.
// aArgVar[] is used for performance; it's assumed to contain valid items only up to aArgIndex, not beyond
// (since normally output vars lie to the left of all input vars, so it doesn't seem worth doing anything
// more complicated).
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
{
	if (mActionType == ACT_SORT) // See PerformSort() for why it's always dereferenced.
		return CONDITION_TRUE;
	aVar = aVar->ResolveAlias(); // Helps performance, but also necessary to accurately detect a match further below.
	if (aVar->Type() == VAR_CLIPBOARD)
		// Even if the clipboard is both an input and an output var, it still
		// doesn't need to be dereferenced into the temp buffer because the
		// clipboard has two buffers of its own.  The only exception is when
		// the clipboard has only files on it, in which case those files need
		// to be converted into plain text:
		return CLIPBOARD_CONTAINS_ONLY_FILES ? CONDITION_TRUE : CONDITION_FALSE;
	if (aVar->Type() != VAR_NORMAL || (!g_NoEnv && !aVar->Length()) || aVar == g_ErrorLevel) // v1.0.43.08: Added g_NoEnv.
		// Reserved vars must always be dereferenced due to their volatile nature.
		// When g_NoEnv==false, normal vars of length zero are dereferenced because they might exist
		// as system environment variables, whose contents are also potentially volatile (i.e. they
		// are sometimes changed by outside forces).
		// As of v1.0.25.12, g_ErrorLevel is always dereferenced also so that a command that sets ErrorLevel
		// can itself use ErrorLevel as in this example: StringReplace, EndKey, ErrorLevel, EndKey:
		return CONDITION_TRUE;

	// Before doing the below, the checks above must be done to ensure it's VAR_NORMAL.  Otherwise, things like
	// the following won't work: StringReplace, o, A_ScriptFullPath, xxx
	// v1.0.45: The following check improves performance slightly by avoiding the loop further below in cases
	// where it's known that a command either doesn't have an output_var or can tolerate the output_var's
	// contents being at the same address as that of one or more of the input-vars.  For example, the commands
	// StringRight/Left and similar can tolerate the same address because they always produce a string whose
	// length is less-than-or-equal to the input-string, thus Assign() will never need to free/realloc the
	// output-var prior to assigning the input-var's contents to it (whose contents are the same as output-var).
	if (!(g_act[mActionType].MaxParamsAu2WithHighBit & 0x80)) // Commands that have this bit don't need final check
		return CONDITION_FALSE;                               // further below (though they do need the ones above).

	// Since the above didn't return, we know that this is a NORMAL input var that isn't an
	// environment variable.  Such input vars only need to be dereferenced if they are also
	// used as an output var by the current script line:
	Var *output_var;
	for (int i = 0; i < mArgc; ++i)
		if (i != aArgIndex && mArg[i].type == ARG_TYPE_OUTPUT_VAR)
		{
			if (   !(output_var = (i < aArgIndex) ? aArgVar[i] : ResolveVarOfArg(i, false))   ) // aArgVar: See top of this function for comments.
				return FAIL;  // It will have already displayed the error.
			if (output_var->ResolveAlias() == aVar)
				return CONDITION_TRUE;
		}
	// Otherwise:
	return CONDITION_FALSE;
}



char *Line::ExpandArg(char *aBuf, int aArgIndex, Var *aArgVar) // 10/2/2006: Doesn't seem worth making it inline due to more complexity than expected.  It would also increase code size without being likely to help performance much.
// Caller must ensure that aArgVar is the variable of the aArgIndex arg when it's of type ARG_TYPE_INPUT_VAR.
// Caller must be sure not to call this for an arg that's marked as an expression, since
// expressions are handled by a different function.  Similarly, it must ensure that none
// of this arg's deref's are function-calls, i.e. that deref->is_function is always false.
// Caller must ensure that aBuf is large enough to accommodate the translation
// of the Arg.  No validation of above params is done, caller must do that.
// Returns a pointer to the char in aBuf that occurs after the zero terminator
// (because that's the position where the caller would normally resume writing
// if there are more args, since the zero terminator must normally be retained
// between args).
{
	ArgStruct &this_arg = mArg[aArgIndex]; // For performance and convenience.
#ifdef _DEBUG
	// This should never be called if the given arg is an output var, so flag that in DEBUG mode:
	if (this_arg.type == ARG_TYPE_OUTPUT_VAR)
	{
		LineError("DEBUG: ExpandArg() was called to expand an arg that contains only an output variable.");
		return NULL;
	}
#endif

	if (aArgVar)
		// +1 so that we return the position after the terminator, as required.
		return aBuf += aArgVar->Get(aBuf) + 1;

	char *pText, *this_marker;
	DerefType *deref;
	for (pText = this_arg.text  // Start at the begining of this arg's text.
		, deref = this_arg.deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
		// Copy the chars that occur prior to deref->marker into the buffer:
		for (this_marker = deref->marker; pText < this_marker; *aBuf++ = *pText++); // this_marker is used to help performance.

		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		aBuf += deref->var->Get(aBuf); // Caller has ensured that deref->is_function==false
		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText; *aBuf++ = *pText++);
	// Terminate the buffer, even if nothing was written into it:
	*aBuf++ = '\0';
	return aBuf; // Returns the position after the terminator.
}



ResultType BackupFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount)
// Helper function for ExpandExpression().  All parameters except the first are output parameters that
// are set for our caller (though caller is responsible for having initialized aVarBackup to NULL).
// If there is nothing to backup, only the aVarBackupCount is changed (to zero).
// Returns OK or FAIL.
{
	if (   !(aVarBackupCount = aFunc.mVarCount + aFunc.mLazyVarCount)   )  // Nothing needs to be backed up.
		return OK; // Leave aVarBackup set to NULL as set by the caller.

	// NOTES ABOUT MALLOC(): Apparently, the implementation of malloc() is quite good, at least for small blocks
	// needed to back up 50 or less variables.  It nearly as fast as alloca(), at least when the system
	// isn't under load and has the memory to spare without swapping.  Therefore, the attempt to use alloca to
	// speed up recursive script-functions didn't result in enough of a speed-up (only 1 to 5%) to be worth the
	// added complexity.
	// Since Var is not a POD struct (it contains private members, a custom constructor, etc.), the VarBkp
	// POD struct is used to hold the backup because it's probably better performance than using Var's
	// constructor to create each backup array element.
	if (   !(aVarBackup = (VarBkp *)malloc(aVarBackupCount * sizeof(VarBkp)))   ) // Caller will take care of freeing it.
		return FAIL;

	int i;
	aVarBackupCount = 0;  // Init only once prior to both loops. aVarBackupCount is being "overloaded" to track the current item in aVarBackup.

	// Note that Backup() does not make the variable empty after backing it up because that is something
	// that must be done by our caller at a later stage.
	for (i = 0; i < aFunc.mVarCount; ++i)
		aFunc.mVar[i]->Backup(aVarBackup[aVarBackupCount++]);
	for (i = 0; i < aFunc.mLazyVarCount; ++i)
		aFunc.mLazyVar[i]->Backup(aVarBackup[aVarBackupCount++]);
	return OK;
}
